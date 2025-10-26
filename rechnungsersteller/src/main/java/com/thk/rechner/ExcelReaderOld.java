package com.thk.rechner;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.InputStream;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * Liest alle Wochentag-Sheets (Mo..So) einer Excel-Datei (.xlsx),
 * sammelt Spielzeiten je Kunde und Tag und gibt sie sortiert in die Konsole aus.
 *
 * Regeln:
 * - Eine Zeit (Std-Belegung) zählt NUR, wenn in DERSELBEN Zeile IRGENDEIN Kundenfeld steht
 *   (z. B. Name, Vorname, Adresse...). Zeilen ohne Kundenfelder gelten als "unbelegt".
 * - Kundendaten werden "nach unten" fortgeschrieben (carry-forward), sobald neue Felder auftauchen.
 * - Zusammenhängende 30-Minuten-Slots werden zu Zeitblöcken gemergt (z. B. 12:00–14:00).
 * - Der "Status" wird ignoriert (nicht benötigt).
 */
public class ExcelReaderOld {

    // ---- Spalten-Namen aus der Kopfzeile (bei dir bitte exakt so) ----
    private static final String COL_HALLE      = "Halle";
    private static final String COL_PLATZ      = "Platz";
    private static final String COL_WOCHENTAG  = "Wochentag";
    private static final String COL_STD_BELEG  = "Std-Belegung";   // Zeit z. B. 12:00, 12:30 ...
    private static final String COL_TARIF      = "Tarif";
    private static final String COL_ANREDE     = "Anrede";
    private static final String COL_TITEL      = "Titel";
    private static final String COL_VORNAME    = "Vorname";
    private static final String COL_NAME       = "Name";
    private static final String COL_ADRESSE    = "Adresse";        // Achtung: In deiner Datei folgt oft noch eine 2. Adressspalte (ohne Header)
    private static final String COL_EMAIL      = "E-Mail";

    // Wochentag-Reihenfolge für sortierte Ausgabe
    private static final List<String> WEEKDAY_ORDER = List.of("Mo","Di","Mi","Do","Fr","Sa","So");

    private static final DateTimeFormatter HHMM = DateTimeFormatter.ofPattern("HH:mm");

    /**
     * Einstieg: Datei lesen, je Kunde & Tag aggregieren und in Konsole ausgeben.
     */
    public static void readAllSheetsGroupByCustomerAndPrint(String excelFilePath) throws Exception {
        try (InputStream in = new FileInputStream(excelFilePath);
             Workbook wb = new XSSFWorkbook(in)) {

            DataFormatter fmt = new DataFormatter(Locale.GERMANY);

            // Kunde -> Aggregation
            Map<String, CustomerAgg> customers = new TreeMap<>();

            for (int si = 0; si < wb.getNumberOfSheets(); si++) {
                Sheet sheet = wb.getSheetAt(si);
                String sheetName = sheet.getSheetName(); // z. B. "Mo"

                Row header = findHeaderRow(sheet);
                if (header == null) continue; // leeres Sheet

                Map<String,Integer> col = mapHeader(header, fmt);

                // Carry: merkt sich die letzten bekannten Kundendaten
                Carry carry = new Carry();

                for (int r = header.getRowNum() + 1; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;

                    // Zeitfeld (Pflicht, sonst ignorieren)
                    String timeStr = get(row, col.get(COL_STD_BELEG), fmt).trim();
                    if (timeStr.isEmpty()) {
                        // Falls die Zeile nur neue Stammdaten bringt (ohne Zeit), Carry aktualisieren
                        if (hasAnyCustomerField(row, col, fmt)) {
                            updateCarry(row, col, fmt, carry);
                        }
                        continue;
                    }

                    // Hat diese ZEILE mindestens ein Kundenfeld? Wenn nein → unbelegt → ignorieren
                    if (!hasAnyCustomerField(row, col, fmt)) {
                        continue;
                    }

                    // Jetzt Carry mit evtl. neuen Werten aus DIESER Zeile aktualisieren
                    updateCarry(row, col, fmt, carry);

                    // Tag bestimmen: Sheetname oder Spalte
                    String weekday = sheetName;
                    String wdFromCell = get(row, col.get(COL_WOCHENTAG), fmt);
                    if (!wdFromCell.isBlank()) weekday = wdFromCell;

                    // Kunde-Key (Eindeutig über Name+Mail, falls vorhanden; ansonsten Name)
                    String customerKey = (carry.name + "|" + carry.vorname + "|" + carry.email).trim();

                    customers.computeIfAbsent(customerKey, k -> new CustomerAgg(carry))
                             .updateStaticInfoFromCarry(carry);

                    // Zeit + weitere Felder sammeln
                    String halle = get(row, col.get(COL_HALLE), fmt);
                    String platz = get(row, col.get(COL_PLATZ), fmt);
                    String tarif = get(row, col.get(COL_TARIF), fmt);

                    LocalTime t = parseTime30(timeStr);
                    if (t == null) continue;

                    customers.get(customerKey)
                             .slots.computeIfAbsent(weekday, d -> new ArrayList<>())
                             .add(new TimeSlot(halle, platz, tarif, t));
                }
            }

            // --- Ausgabe: je Kunde ---
            for (CustomerAgg agg : customers.values()) {
                String nameZeile = String.join(" ",
                    orEmpty(agg.anrede),
                    orEmpty(agg.titel),
                    orEmpty(agg.vorname),
                    orEmpty(agg.name)
                ).replaceAll("\\s{2,}", " ").trim();

                System.out.println((nameZeile.isEmpty() ? "" : nameZeile + " :"));
                if (notBlank(agg.adresse)) System.out.println(agg.adresse);
                if (notBlank(agg.email))   System.out.println(agg.email);
                System.out.println("Spielzeit :");
                System.out.println();

                // --- Tage in fester Reihenfolge (Mo..So) ---
                for (String day : WEEKDAY_ORDER) {
                    List<TimeSlot> slots = agg.slots.get(day);
                    if (slots == null || slots.isEmpty()) continue;

                    // 1) sortieren
                    slots.sort(Comparator
                            .comparing((TimeSlot ts) -> safe(ts.halle))
                            .thenComparing(ts -> safe(ts.platz))
                            .thenComparing(ts -> safe(ts.tarif))
                            .thenComparing(ts -> ts.time));

                    // 2) gruppieren nach (Halle, Platz, Tarif) und innerhalb gruppieren zu Blöcken
                    List<SlotBlock> blocks = compressByCourtAndTariff(day, slots);

                    // 3) ausgeben
                    for (SlotBlock b : blocks) {
                        String halle = notBlank(b.halle) ? "Halle " + b.halle : "";
                        String platz = notBlank(b.platz) ? "Platz " + b.platz : "";

                        String line = String.join("\t",
                                halle,
                                platz,
                                b.day,
                                HHMM.format(b.start) + " - " + HHMM.format(b.end),
                                orEmpty(b.tarif)
                        ).replaceAll("\\t{2,}", "\t");

                        System.out.println(line);
                    }
                    System.out.println(); // Leerzeile zwischen Tagen
                }
                System.out.println(); // Leerzeile zwischen Kunden
            }
        }
    }

    // ----- Modelle -----

    private static class Carry {
        String anrede = "";
        String titel = "";
        String vorname = "";
        String name = "";
        String adresse = "";
        String email = "";
    }

    private static class TimeSlot {
        final String halle;
        final String platz;
        final String tarif;
        final LocalTime time;

        TimeSlot(String halle, String platz, String tarif, LocalTime time) {
            this.halle = orEmpty(halle);
            this.platz = orEmpty(platz);
            this.tarif = orEmpty(tarif);
            this.time  = time;
        }
    }

    private static class SlotBlock {
        final String day, halle, platz, tarif;
        final LocalTime start, end;

        SlotBlock(String day, String halle, String platz, String tarif, LocalTime start, LocalTime end) {
            this.day = day;
            this.halle = halle;
            this.platz = platz;
            this.tarif = tarif;
            this.start = start;
            this.end = end;
        }
    }

    private static class CustomerAgg {
        String anrede, titel, vorname, name, adresse, email;
        Map<String, List<TimeSlot>> slots = new TreeMap<>(weekdayComparator());

        CustomerAgg(Carry c) { updateStaticInfoFromCarry(c); }

        void updateStaticInfoFromCarry(Carry c) {
            if (notBlank(c.anrede))  this.anrede  = c.anrede;
            if (notBlank(c.titel))   this.titel   = c.titel;
            if (notBlank(c.vorname)) this.vorname = c.vorname;
            if (notBlank(c.name))    this.name    = c.name;
            if (notBlank(c.adresse)) this.adresse = c.adresse;
            if (notBlank(c.email))   this.email   = c.email;
        }
    }

    private static Comparator<String> weekdayComparator() {
        return (a, b) -> {
            int ia = WEEKDAY_ORDER.indexOf(a);
            int ib = WEEKDAY_ORDER.indexOf(b);
            if (ia == -1 && ib == -1) return a.compareTo(b);
            if (ia == -1) return 1;
            if (ib == -1) return -1;
            return Integer.compare(ia, ib);
        };
    }

    // ----- Header/Parsing -----

    /** Sucht die erste sinnvolle Header-Zeile (erste Zeile mit mindestens einer nicht-leeren Zelle) */
    private static Row findHeaderRow(Sheet sheet) {
        DataFormatter fmt = new DataFormatter(Locale.GERMANY);
        for (Row row : sheet) {
            for (Cell c : row) {
                String val = (c == null) ? "" : fmt.formatCellValue(c);
                if (!val.isBlank()) {
                    return row;
                }
            }
        }
        return null;
    }

    /** Mappt Headernamen -> Spaltenindex (Case-insensitive; trim) */
    private static Map<String,Integer> mapHeader(Row header, DataFormatter fmt) {
        Map<String,Integer> map = new HashMap<>();
        for (Cell c : header) {
            if (c == null) continue;
            String name = fmt.formatCellValue(c).trim();
            if (name.isEmpty()) continue;
            map.put(name, c.getColumnIndex());
        }
        // Sicherstellen, dass wichtige Spalten da sind (sonst nicht schlimm, wird übersprungen)
        return map;
    }

    /** Liefert die Zelle selbst oder – falls sie leer ist und in einem Merge-Bereich liegt –
     *  die linke obere Zelle des Merge-Bereichs. Damit werden in Excel zusammengefasste
     *  Zellen korrekt gelesen (häufige Ursache für „verrutschte“ Anrede/Name). */
    private static Cell getEffectiveCell(Row row, int colIdx) {
        if (row == null || colIdx < 0) return null;
        Sheet sheet = row.getSheet();
        Cell cell = row.getCell(colIdx);
        if (cell != null && cell.getCellType() != CellType.BLANK &&
            !(cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty())) {
            return cell;
        }
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.isInRange(row.getRowNum(), colIdx)) {
                Row topRow = sheet.getRow(region.getFirstRow());
                if (topRow == null) return null;
                return topRow.getCell(region.getFirstColumn());
            }
        }
        return cell;
    }

    /** Holt String-Inhalt einer Zelle (leer = ""), merge-sicher */
    private static String get(Row row, Integer colIdx, DataFormatter fmt) {
        if (colIdx == null) return "";
        Cell eff = getEffectiveCell(row, colIdx);
        return (eff == null) ? "" : fmt.formatCellValue(eff).trim();
    }

    /** Prüft, ob in der Zeile irgendein Kundenfeld gesetzt ist (Anrede/Titel/Vorname/Name/Adresse/Email) */
    private static boolean hasAnyCustomerField(Row row, Map<String,Integer> col, DataFormatter fmt) {
        if (notBlank(get(row, col.get(COL_ANREDE),  fmt))) return true;
        if (notBlank(get(row, col.get(COL_TITEL),   fmt))) return true;
        if (notBlank(get(row, col.get(COL_VORNAME), fmt))) return true;
        if (notBlank(get(row, col.get(COL_NAME),    fmt))) return true;

        // Adresse: in deiner Datei oft auf ZWEI Spalten verteilt (Adresse + nächste leere Überschrift).
        if (notBlank(get(row, col.get(COL_ADRESSE), fmt))) return true;

        // Versuch: zweite (rechte) Adresshälfte -> direkt die Spalte rechts neben "Adresse"
        if (col.get(COL_ADRESSE) != null) {
            String addr2 = getByIndex(row, col.get(COL_ADRESSE) + 1, fmt);
            if (notBlank(addr2)) return true;
        }
        if (notBlank(get(row, col.get(COL_EMAIL), fmt))) return true;

        return false;
    }

    /** Holt String anhand eines konkreten Spaltenindex (auch wenn Header leer ist), merge-sicher */
    private static String getByIndex(Row row, int colIndex, DataFormatter fmt) {
        if (colIndex < 0) return "";
        Cell eff = getEffectiveCell(row, colIndex);
        return (eff == null) ? "" : fmt.formatCellValue(eff).trim();
    }

    /** Aktualisiert den Carry mit allen vorhandenen Kundendaten aus dieser Zeile (leere Werte überschreiben NICHT) */
    private static void updateCarry(Row row, Map<String,Integer> col, DataFormatter fmt, Carry carry) {
        carry.anrede  = keepOrUpdate(carry.anrede,  get(row, col.get(COL_ANREDE),  fmt));
        carry.titel   = keepOrUpdate(carry.titel,   get(row, col.get(COL_TITEL),   fmt));
        carry.vorname = keepOrUpdate(carry.vorname, get(row, col.get(COL_VORNAME), fmt));
        carry.name    = keepOrUpdate(carry.name,    get(row, col.get(COL_NAME),    fmt));

        // Adresse: linke Spalte + optional rechte Nachbarspalte (falls vorhanden) zusammenführen
        String addr1 = get(row, col.get(COL_ADRESSE), fmt);
        String addr2 = "";
        if (col.get(COL_ADRESSE) != null) {
            addr2 = getByIndex(row, col.get(COL_ADRESSE) + 1, fmt);
        }
        String mergedAddr = mergeAddress(addr1, addr2);
        carry.adresse = keepOrUpdate(carry.adresse, mergedAddr);

        carry.email   = keepOrUpdate(carry.email,   get(row, col.get(COL_EMAIL),   fmt));
    }

    /** Nimmt neuen Wert nur dann, wenn er nicht leer ist; sonst bleibt der alte erhalten */
    private static String keepOrUpdate(String oldVal, String newVal) {
        String n = orEmpty(newVal).trim();
        if (!n.isEmpty()) return n;
        return orEmpty(oldVal);
    }

    /** Adresshälften sauber zusammenführen */
    private static String mergeAddress(String a, String b) {
        String s1 = orEmpty(a).trim();
        String s2 = orEmpty(b).trim();
        if (s1.isEmpty() && s2.isEmpty()) return "";
        if (s1.isEmpty()) return s2;
        if (s2.isEmpty()) return s1;
        String s = (s1 + " " + s2).replaceAll("\\s{2,}", " ").trim();
        // Optional: häufiges Muster "Straße + PLZ Ort" -> in 2 Zeilen:
        // Erkennen durch 5-stellige PLZ
        if (s.matches(".*\\b\\d{5}\\b.*")) {
            // Versuche, vor der PLZ einen Zeilenumbruch zu setzen
            return s.replaceFirst("\\s(\\d{5})\\b", "\n$1");
        }
        return s;
    }

    private static String orEmpty(String s) { return s == null ? "" : s; }
    private static boolean notBlank(String s) { return s != null && !s.trim().isEmpty(); }

    // ----- Zeit-Parsing & Komprimierung -----

    /** Erwartet 30-Minuten-Raster (12:00, 12:30, ...). Gibt null zurück, wenn unparsebar. */
    private static LocalTime parseTime30(String t) {
        String s = orEmpty(t).trim();
        if (!s.matches("\\d{1,2}:\\d{2}")) return null;
        try {
            LocalTime time = LocalTime.parse(s, HHMM);
            int min = time.getMinute();
            if (min == 0 || min == 30) return time;
        } catch (Exception ignore) {}
        return null;
    }

    /** Komprimiert Slots pro (Halle, Platz, Tarif) zu Blöcken (start..end, end exklusiv) */
    private static List<SlotBlock> compressByCourtAndTariff(String day, List<TimeSlot> in) {
        List<SlotBlock> res = new ArrayList<>();
        if (in.isEmpty()) return res;

        String curHalle = null, curPlatz = null, curTarif = null;
        LocalTime blockStart = null;
        LocalTime prevTime = null;

        for (TimeSlot s : in) {
            boolean sameGroup = Objects.equals(curHalle, s.halle)
                    && Objects.equals(curPlatz, s.platz)
                    && Objects.equals(curTarif, s.tarif);

            if (curHalle == null) {
                // erster Slot
                curHalle = s.halle; curPlatz = s.platz; curTarif = s.tarif;
                blockStart = s.time;
                prevTime = s.time;
                continue;
            }

            if (sameGroup && s.time.equals(prevTime.plusMinutes(30))) {
                // direkt anschließender Slot -> Block erweitern
                prevTime = s.time;
            } else {
                // Block schließen und neu beginnen
                res.add(new SlotBlock(day, curHalle, curPlatz, curTarif, blockStart, prevTime.plusMinutes(30)));
                curHalle = s.halle; curPlatz = s.platz; curTarif = s.tarif;
                blockStart = s.time;
                prevTime = s.time;
            }
        }

        // letzten Block schließen
        if (curHalle != null) {
            res.add(new SlotBlock(day, curHalle, curPlatz, curTarif, blockStart, prevTime.plusMinutes(30)));
        }

        return res;
    }

    private static String safe(String s) { return s == null ? "" : s.trim(); }

}
