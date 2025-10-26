package com.thk.rechner;

import org.apache.poi.xwpf.usermodel.*;

import java.awt.Desktop;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

public final class InvoiceGenerator {

    private InvoiceGenerator() {}

    // --- Blöcke (zusammenhängende 30-Minuten-Slots) ---
    private static final class Block {
        final String sheet;     // Mo/Di/…
        final String hall;      // Halle
        final String court;     // Platz
        final String start;     // HH:mm
        final String end;       // HH:mm
        final int units30;      // Anzahl 30-Min-Einheiten
        Block(String sheet, String hall, String court, String start, String end, int units30) {
            this.sheet = sheet; this.hall = hall; this.court = court;
            this.start = start; this.end = end; this.units30 = units30;
        }
    }

    public static Path generate(ExcelReader.CustomerData data, Path templatePath) throws IOException {
        Objects.requireNonNull(data, "data must not be null");
        Objects.requireNonNull(templatePath, "templatePath must not be null");

        // --- 1) Datum & Adresse
        LocalDate today = LocalDate.now();
        String datum = today.format(DateTimeFormatter.ofPattern("dd.MM.yyyy"));
        String jahr  = String.valueOf(today.getYear());

        String adresse = data.address();
        String plzUndStadt = null;
        
        if (adresse != null) {
            String trimmed = adresse.trim();
            java.util.regex.Matcher m = java.util.regex.Pattern
                    .compile("(?i)(?:D\\s*-\\s*)?(\\d{5})")
                    .matcher(trimmed);
            if (m.find()) {
                int idx = m.start(1); // Start der eigentlichen 5 Ziffern
                plzUndStadt = trimmed.substring(idx).trim();
                adresse = trimmed.substring(0, idx).trim(); // kann leer sein, ist ok
            }
        }

        // --- 2) Preis je 30 Minuten
        Double pricePerUnit = null;
        for (ExcelReader.BookingEntry b : data.bookings()) {
            if (b.price() != null) { pricePerUnit = b.price(); break; }
        }

        // --- 3) Blöcke bilden
        final int SLOT_MINUTES = 30;
        Map<String, List<Integer>> startsByKey = new LinkedHashMap<>();
        for (ExcelReader.BookingEntry b : data.bookings()) {
            String from = prefer(b.timeFrom(), fromRaw(b.timeRaw()));
            if (from == null || from.isBlank()) continue;
            int m = parseHmToMinutes(from);
            if (m < 0) continue;
            String key = (nz(b.sheet())+"|"+nz(b.hall())+"|"+nz(b.court()));
            startsByKey.computeIfAbsent(key, k -> new ArrayList<>()).add(m);
        }
        List<Block> blocks = new ArrayList<>();
        for (Map.Entry<String,List<Integer>> e : startsByKey.entrySet()) {
            String[] parts = e.getKey().split("\\|", -1);
            String sheet = parts.length>0?parts[0]:"";
            String hall  = parts.length>1?parts[1]:"";
            String court = parts.length>2?parts[2]:"";
            List<Integer> starts = e.getValue();
            Collections.sort(starts);
            int segStart = -1, last = -1;
            for (int s : starts) {
                if (segStart < 0) { segStart = s; last = s; continue; }
                if (s == last + SLOT_MINUTES) { last = s; continue; }
                blocks.add(new Block(
                        sheet, hall, court,
                        formatMinutesAsHHmm(segStart),
                        formatMinutesAsHHmm(last + SLOT_MINUTES),
                        (last - segStart + SLOT_MINUTES) / SLOT_MINUTES
                ));
                segStart = s; last = s;
            }
            if (segStart >= 0) {
                blocks.add(new Block(
                        sheet, hall, court,
                        formatMinutesAsHHmm(segStart),
                        formatMinutesAsHHmm(last + SLOT_MINUTES),
                        (last - segStart + SLOT_MINUTES) / SLOT_MINUTES
                ));
            }
        }
        if (blocks.isEmpty() && !data.bookings().isEmpty()) {
            ExcelReader.BookingEntry a = data.bookings().get(0);
            String from = prefer(a.timeFrom(), fromRaw(a.timeRaw()));
            String to   = prefer(a.timeTo(),   toRaw(a.timeRaw()));
            int units = 0;
            int fm = parseHmToMinutes(from);
            int tm = parseHmToMinutes(to);
            if (fm >= 0 && tm > fm) units = (tm - fm) / SLOT_MINUTES;
            blocks.add(new Block(nz(a.sheet()), nz(a.hall()), nz(a.court()),
                    cleanHm(from), cleanHm(to), Math.max(units, 1)));
        }

        // --- 4) Summen
        int totalUnits30 = blocks.stream().mapToInt(b -> b.units30).sum();
        Double brutto = (pricePerUnit == null) ? null : pricePerUnit * totalUnits30;
        String bruttoSumme = (brutto == null) ? "" : formatPrice(brutto);

        // --- 5) DOCX laden
        XWPFDocument doc;
        try (InputStream in = Files.newInputStream(templatePath)) {
            doc = new XWPFDocument(in);
        }

        // --- 6) Tabellenzeile mit {{spieltag}} finden (Layout-/Datenzeile)
        XWPFTable targetTable = null;
        XWPFTableRow targetRow = null;
        int targetRowIndex = -1;
        searchLoop:
        for (XWPFTable t : doc.getTables()) {
            List<XWPFTableRow> rows = t.getRows();
            for (int i = 0; i < rows.size(); i++) {
                XWPFTableRow r = rows.get(i);
                String txt = getRowText(r);
                if (txt != null && txt.contains("{{spieltag}}")) {
                    targetTable = t;
                    targetRow = r;
                    targetRowIndex = i;
                    break searchLoop;
                }
            }
        }

        // --- 7) Zielzeile befüllen und für weitere Blöcke neue Zeilen erzeugen
        if (targetTable != null && targetRow != null && !blocks.isEmpty()) {
            // sicherstellen, dass 3 Zellen existieren
            while (targetRow.getTableCells().size() < 3) targetRow.addNewTableCell();

            // Block #1 in der vorhandenen Zeile
            Block b0 = blocks.get(0);
            fillRowWithBlock(targetRow, b0, pricePerUnit, SLOT_MINUTES);

            // weitere Blöcke: jeweils neue Zeile UNTERHALB einfügen
            int inserted = 0;
            for (int i = 1; i < blocks.size(); i++) {
                Block b = blocks.get(i);
                XWPFTableRow newRow = insertBlankRowLike(targetTable, targetRow, targetRowIndex + inserted + 1);
                fillRowWithBlock(newRow, b, pricePerUnit, SLOT_MINUTES);
                inserted++;
            }
        }

        // --- 8) Platzhalter außerhalb der dynamischen Zelle
        Map<String, String> vars = new HashMap<>();
        vars.put("anrede",        nz(data.salutation()));
        if (data.title() != null) {
        	vars.put("titel",         nz(" " + data.title()));
        } else {
        	vars.put("titel",         nz(data.title()));
        }
        vars.put("vorname",       nz(data.firstName()));
        vars.put("name",          nz(data.lastName()));
        vars.put("adresse",       nz(adresse));
        vars.put("plzundstadt",   nz(plzUndStadt));
        vars.put("datum",         datum);
        vars.put("jahr",          jahr);
        vars.put("BruttoSumme",   nz(bruttoSumme));
        
        double netto = brutto / 1.19;
        double ust = brutto - netto;

        // Neue Platzhalter ergänzen
        vars.put("netto", String.format("%.2f €", netto));
        vars.put("erhalteneUst", String.format("%.2f €", ust));

        replaceAll(doc, vars);

        // --- 9) Schreiben
        Path outPath = buildOutputPath(data.lastName(), today);
        try (OutputStream out = Files.newOutputStream(outPath, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING)) {
            doc.write(out);
        }
        doc.close();

       
        return outPath;
    }

    // --- Tabellen-Helfer ---

    /** Fügt unterhalb von refRow eine neue leere Zeile mit gleicher Zellanzahl und Formatierung ein. */
    private static XWPFTableRow insertBlankRowLike(XWPFTable table, XWPFTableRow refRow, int insertPos) {
        // Neue Zeile an Position erzeugen
        XWPFTableRow newRow = table.insertNewTableRow(insertPos);
        
        // Zellenformatierung von der Referenzzeile kopieren
        for (int i = 0; i < refRow.getTableCells().size(); i++) {
            XWPFTableCell refCell = refRow.getCell(i);
            XWPFTableCell newCell = newRow.addNewTableCell();
            
            // Zellenformatierung kopieren
            if (refCell != null) {
                try {
                    // TcPr für neue Zelle sicher erstellen
                    if (newCell.getCTTc().getTcPr() == null) {
                        newCell.getCTTc().addNewTcPr();
                    }
                    
                    // Referenz-TcPr kopieren falls vorhanden
                    if (refCell.getCTTc().getTcPr() != null) {
                        // GridSpan kopieren
                        if (refCell.getCTTc().getTcPr().getGridSpan() != null) {
                            newCell.getCTTc().getTcPr().addNewGridSpan().setVal(refCell.getCTTc().getTcPr().getGridSpan().getVal());
                        }
                        
                        // Vertikale Ausrichtung kopieren
                        if (refCell.getCTTc().getTcPr().getVAlign() != null) {
                            newCell.getCTTc().getTcPr().addNewVAlign().setVal(refCell.getCTTc().getTcPr().getVAlign().getVal());
                        }
                    }
                    
                    // Absatzformatierung kopieren
                    if (!refCell.getParagraphs().isEmpty()) {
                        XWPFParagraph refPara = refCell.getParagraphs().get(0);
                        XWPFParagraph newPara = newCell.getParagraphs().get(0);
                        
                        // Ausrichtung kopieren
                        if (refPara.getAlignment() != null) {
                            newPara.setAlignment(refPara.getAlignment());
                        }
                    }
                } catch (Exception e) {
                    // Falls Fehler auftreten, einfach mit Standardformatierung fortfahren
                    System.err.println("Warnung: Fehler beim Kopieren der Zellenformatierung: " + e.getMessage());
                }
            }
            
            // Text mit Standardformatierung setzen
            clearAndSetCellText(newCell, "");
        }
        return newRow;
    }

    /** Schreibt Beschreibung / Stunden / Betrag in die 3 Zellen der Zeile. */
    /** Schreibt Beschreibung / Stunden / Betrag in die 3 Zellen der Zeile. */
    private static void fillRowWithBlock(XWPFTableRow row, Block b, Double pricePerUnit, int SLOT_MINUTES) {
        while (row.getTableCells().size() < 3) row.addNewTableCell();

        String descTemplate = "{{spieltag}} Platz {{platz}} von {{startzeit}} bis {{endzeit}} Uhr";
        Map<String,String> local = Map.of(
                "spieltag", nz(b.sheet),
                "platz",    nz(b.court),
                "startzeit",nz(b.start),
                "endzeit",  nz(b.end)
        );
        
        // Zelle 0: Beschreibung (mit Platzhaltern, falls im Template vorhanden)
        // Statt clearAndSetCellText verwenden wir setCellFromTemplate um Formatierung zu erhalten
        setCellFromTemplate(row.getCell(0), descTemplate, local,
                nz(b.sheet) + " Platz " + nz(b.court) + " von " + nz(b.start) + " bis " + nz(b.end) + " Uhr");

        // Zelle 1: Leistung/Std - Formatierung der Ursprungszelle beibehalten
        double hours = (b.units30 * SLOT_MINUTES) / 60.0;
        String hoursText = formatHoursNumber(hours);
        XWPFTableCell hoursCell = row.getCell(1);
        if (hoursCell.getParagraphs().isEmpty()) {
            clearAndSetCellText(hoursCell, hoursText);
        } else {
            // Bestehende Formatierung beibehalten, nur Text ersetzen
            clearAndSetCellText(hoursCell, hoursText);
        }

        // Zelle 2: Betrag brutto - Formatierung der Ursprungszelle beibehalten
        String priceText = (pricePerUnit == null) ? "" : formatPrice(pricePerUnit * b.units30);
        XWPFTableCell priceCell = row.getCell(2);
        if (priceCell.getParagraphs().isEmpty()) {
            clearAndSetCellText(priceCell, priceText);
        } else {
            // Bestehende Formatierung beibehalten, nur Text ersetzen
            clearAndSetCellText(priceCell, priceText);
        }
    }

    private static String getRowText(XWPFTableRow row) {
        StringBuilder sb = new StringBuilder();
        for (XWPFTableCell c : row.getTableCells()) {
            for (XWPFParagraph p : c.getParagraphs()) {
                for (XWPFRun r : p.getRuns()) {
                    sb.append(r.toString());
                }
            }
        }
        return sb.toString();
    }

    /** Zelle komplett leeren und einfachen Text setzen (ein Absatz, ein Run). */
    private static void clearAndSetCellText(XWPFTableCell cell, String text) {
        // Ausrichtung der ersten Zelle speichern
        ParagraphAlignment alignment = null;
        if (!cell.getParagraphs().isEmpty()) {
            alignment = cell.getParagraphs().get(0).getAlignment();
        }
        
        // Alle Absätze entfernen
        int cnt = cell.getParagraphs().size();
        for (int i = cnt - 1; i >= 0; i--) {
            cell.removeParagraph(i);
        }
        
        // Neuen Absatz mit Text erstellen
        XWPFParagraph p = cell.addParagraph();
        if (alignment != null) {
            p.setAlignment(alignment);
        }
        XWPFRun run = p.createRun();
        run.setFontFamily("Arial");
        run.setFontSize(10);
        run.setText(nz(text));
    }

    /** Ersetzt {{keys}} im übergebenen templateText mit localVars; wenn keine Platzhalter, setzt fallbackText. */
    private static void setCellFromTemplate(XWPFTableCell cell, String templateText, Map<String,String> localVars, String fallbackText) {
        String existing = cell.getText();
        boolean useTemplate = existing != null && existing.contains("{{");

        // Zelle leeren
        int cnt = cell.getParagraphs().size();
        for (int i = cnt - 1; i >= 0; i--) cell.removeParagraph(i);

        XWPFParagraph p = cell.addParagraph();
        String source = useTemplate ? templateText : fallbackText;

        java.util.regex.Pattern pat = java.util.regex.Pattern.compile("\\{\\{([a-zA-Z0-9_]+)\\}\\}");
        java.util.regex.Matcher m = pat.matcher(source);
        int last = 0;
        while (m.find()) {
            if (m.start() > last) {
                XWPFRun r = p.createRun(); r.setFontFamily("Arial"); r.setFontSize(10);
                r.setText(source.substring(last, m.start()));
            }
            String key = m.group(1);
            String val = localVars.getOrDefault(key, "");
            XWPFRun r = p.createRun(); r.setFontFamily("Arial"); r.setFontSize(10); r.setBold(true);
            r.setText(val);
            last = m.end();
        }
        if (last < source.length()) {
            XWPFRun r = p.createRun(); r.setFontFamily("Arial"); r.setFontSize(10);
            r.setText(source.substring(last));
        }
    }

    // --- Platzhalter-Replacement für restliches Dokument ---

    private static void replaceAll(XWPFDocument doc, Map<String, String> vars) {
        // Absätze
        for (XWPFParagraph p : doc.getParagraphs()) {
            replaceInParagraph(p, vars);
        }
        // Tabellen
        for (XWPFTable t : doc.getTables()) {
            for (XWPFTableRow row : t.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        replaceInParagraph(p, vars);
                    }
                }
            }
        }
        // Kopf-/Fußzeilen
        for (XWPFHeader h : doc.getHeaderList()) {
            for (XWPFParagraph p : h.getParagraphs()) replaceInParagraph(p, vars);
            for (XWPFTable t : h.getTables()) {
                for (XWPFTableRow r : t.getRows())
                    for (XWPFTableCell c : r.getTableCells())
                        for (XWPFParagraph p : c.getParagraphs()) replaceInParagraph(p, vars);
            }
        }
        for (XWPFFooter f : doc.getFooterList()) {
            for (XWPFParagraph p : f.getParagraphs()) replaceInParagraph(p, vars);
            for (XWPFTable t : f.getTables()) {
                for (XWPFTableRow r : t.getRows())
                    for (XWPFTableCell c : r.getTableCells())
                        for (XWPFParagraph p : c.getParagraphs()) replaceInParagraph(p, vars);
            }
        }
    }

    private static void replaceInParagraph(XWPFParagraph paragraph, Map<String, String> vars) {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs == null || runs.isEmpty()) return;

        StringBuilder sb = new StringBuilder();
        for (XWPFRun r : runs) sb.append(r.toString());
        String original = sb.toString();
        if (original.isEmpty()) return;

        java.util.regex.Pattern pat = java.util.regex.Pattern.compile("\\{\\{([a-zA-Z0-9_]+)\\}\\}");
        java.util.regex.Matcher mCheck = pat.matcher(original);
        if (!mCheck.find()) return;

        // Bestehende Runs entfernen
        for (int i = runs.size() - 1; i >= 0; i--) paragraph.removeRun(i);

        java.util.regex.Matcher m = pat.matcher(original);
        int last = 0;

        while (m.find()) {
            // Text VOR dem Platzhalter ausgeben (Standardformat)
            if (m.start() > last) {
                String text = original.substring(last, m.start());
                if (!text.isEmpty()) {
                    XWPFRun run = paragraph.createRun();
                    run.setFontFamily("Arial");
                    run.setFontSize(10);
                    run.setBold(false);
                    run.setItalic(false);
                    run.setText(text);
                }
            }

            // Platzhalter ersetzen und formatiert ausgeben
            String key = m.group(1);
            String val = vars.get(key);
            if (val == null) val = "";

            XWPFRun run = paragraph.createRun();
            run.setFontFamily("Arial");

            // Formatregeln für bestimmte Platzhalter
            if ("BruttoSumme".equals(key)) {
                run.setFontSize(11);
                run.setBold(true);
                run.setItalic(false);
            } else if ("erhalteneUst".equalsIgnoreCase(key) || "netto".equalsIgnoreCase(key)) {
                run.setFontSize(8);
                run.setBold(false);
                run.setItalic(true);
            } else {
                // alle übrigen Platzhalter: Standard
                run.setFontSize(10);
                run.setBold(false);
                run.setItalic(false);
            }

            run.setText(val);
            last = m.end();
        }

        // Restlicher Text NACH dem letzten Platzhalter (Standardformat)
        if (last < original.length()) {
            String tail = original.substring(last);
            if (!tail.isEmpty()) {
                XWPFRun run = paragraph.createRun();
                run.setFontFamily("Arial");
                run.setFontSize(10);
                run.setBold(false);
                run.setItalic(false);
                run.setText(tail);
            }
        }
    }

    // --- Utilities ---

    private static String prefer(String a, String b) { return (a != null && !a.isBlank()) ? a : b; }
    private static String nz(String s) { return s == null ? "" : s; }

    private static String fromRaw(String raw) {
        if (raw == null) return null;
        String[] ft = splitTimeRange(raw);
        return ft[0];
    }
    private static String toRaw(String raw) {
        if (raw == null) return null;
        String[] ft = splitTimeRange(raw);
        return ft[1];
    }

    private static int parseHmToMinutes(String hm) {
        if (hm == null || hm.isBlank()) return -1;
        String s = hm.trim();
        if (!s.matches("^\\d{1,2}:\\d{2}$")) return -1;
        String[] p = s.split(":");
        try {
            int h = Integer.parseInt(p[0]);
            int m = Integer.parseInt(p[1]);
            return h * 60 + m;
        } catch (Exception e) {
            return -1;
        }
    }

    private static String formatHoursNumber(double hours) {
        java.text.NumberFormat nf = java.text.NumberFormat.getNumberInstance(java.util.Locale.GERMANY);
        nf.setMinimumFractionDigits(0);
        nf.setMaximumFractionDigits(2);
        String s = nf.format(hours);
        if (s.endsWith(",00")) s = s.substring(0, s.length() - 3);
        if (s.endsWith(",0"))  s = s.substring(0, s.length() - 2);
        return s;
    }

    private static String formatMinutesAsHHmm(int minutes) {
        if (minutes < 0) return "";
        int h = minutes / 60;
        int m = minutes % 60;
        return String.format("%02d:%02d", h, m);
    }

    private static String formatPrice(double v) {
        return String.format(Locale.GERMANY, "%,.2f €", v);
    }

    private static int indexOfRegex(String input, String regex) {
        var m = java.util.regex.Pattern.compile(regex).matcher(input);
        return m.find() ? m.start() : -1;
    }

    private static Path buildOutputPath(String lastName, LocalDate today) throws IOException {
        Path baseDir = Paths.get(System.getProperty("user.home"), "Desktop", "Rechnungen");
        String folderName = today.format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
        Path outDir = baseDir.resolve(folderName);
        Files.createDirectories(outDir);

        String safeLastName = (lastName == null || lastName.isBlank())
                ? "Unbekannt"
                : lastName.replaceAll("[^\\p{L}\\p{Nd}_-]", "");
        return outDir.resolve("THK_xxx-2025_" + safeLastName + ".docx");
    }

    /** "12:00 - 14:00" → ["12:00","14:00"] */
    private static String[] splitTimeRange(String timeRaw) {
        if (timeRaw == null) return new String[]{null, null};
        String norm = timeRaw.replace('\u2013', '-').replace('\u2014', '-').trim();
        String[] parts = norm.split("\\s*-\\s*");
        String from = parts.length > 0 ? cleanHm(parts[0]) : null;
        String to   = parts.length > 1 ? cleanHm(parts[1]) : null;
        return new String[]{from, to};
    }

    private static String cleanHm(String s) {
        if (s == null) return null;
        String t = s.trim();
        if (t.matches("^\\d{1,2}:\\d{2}:\\d{2}$")) return t.substring(0,5);
        if (t.matches("^\\d{1,2}:\\d{2}$")) return t;
        return null;
    }
}
