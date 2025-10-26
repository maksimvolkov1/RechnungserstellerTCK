package com.thk.rechner;

import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

/**
 * Excel-Reader für den Rechnungsersteller.
 * Unterstützt mehrere Sheets ("Mo", "Di", "Mi", "Do", "Fr", "Sa", "So").
 *
 * Bietet u.a.:
 *  - groupRowsByCustomerAllSheets(...)  → Kunde → Fundstellen
 *  - readCustomer(excelPath, customer)  → Alle Informationen eines Kunden (Kontakt + Buchungen)
 */
public final class ExcelReader {

    /** Standard-Sheetnamen (deutsche Wochentage, abgekürzt) */
    public static final List<String> DEFAULT_SHEETS =
            List.of("Mo", "Di", "Mi", "Do", "Fr", "Sa", "So");

    private ExcelReader() {}

    /** Sheet + 1-basierte Zeilennummer */
    public record RowRef(String sheetName, int rowNumber) {
        @Override public String toString() { return sheetName + "!" + rowNumber; }
    }

    /** Einzelne Buchungszeile des Kunden */
    public record BookingEntry(
            String sheet,           // z. B. "Mo"
            int rowNumber,          // 1-basiert (wie Excel)
            String hall,            // "Halle 1" oder "1"
            String court,           // "Platz 2" oder "2"
            String timeFrom,        // "12:00"
            String timeTo,          // "14:00"
            String timeRaw,         // Original-Zeitwert (falls nur ein Feld "12:00 - 14:00")
            Double price            // 290.00
    ) {}

    /** Kundendaten + alle Buchungen */
    public record CustomerData(
            String salutation,      // Anrede
            String title,           // Titel
            String firstName,       // Vorname
            String lastName,        // Name/Nachname
            String email,           // E-Mail
            String address,         // Adresse (+ evtl. PLZ/Ort zusammengeführt)
            List<BookingEntry> bookings
    ) {}

    // ------------------------------------------------------------
    // Gruppierung nach Kunden
    // ------------------------------------------------------------
    public static Map<String, List<RowRef>> groupRowsByCustomerAllSheets(Path excelPath) throws IOException {
        return groupRowsByCustomer(excelPath, DEFAULT_SHEETS);
    }

    public static Map<String, List<RowRef>> groupRowsByCustomer(Path excelPath, Collection<String> sheetNames)
            throws IOException {

        Objects.requireNonNull(excelPath, "excelPath must not be null");
        Objects.requireNonNull(sheetNames, "sheetNames must not be null");

        Map<String, List<RowRef>> result = new LinkedHashMap<>();

        try (InputStream in = Files.newInputStream(excelPath);
             Workbook wb = WorkbookFactory.create(in)) {

            for (String desired : sheetNames) {
                Sheet sheet = wb.getSheet(desired);
                if (sheet == null) continue;

                Row header = findFirstNonEmptyRow(sheet);
                if (header == null) continue;

                int nameCol = findColumnIndex(header, List.of("name"));
                if (nameCol < 0) continue;

                for (int r = header.getRowNum() + 1; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;

                    String customerName = getCellString(row.getCell(nameCol));
                    if (isBlank(customerName)) continue;

                    result.computeIfAbsent(customerName.trim(), k -> new ArrayList<>())
                          .add(new RowRef(desired, r + 1));
                }
            }
        }

        return result;
    }

    // ------------------------------------------------------------
    // Alle Infos für einen Kunden einlesen
    // ------------------------------------------------------------
    /**
     * Liest über alle Wochentag-Sheets alle Zeilen eines Kunden und gibt strukturierte Daten zurück.
     * Erwartete (optionale) Kopfspalten – case-insensitive:
     *  - "Name" (Pflicht für die Zuordnung), "Vorname", "Anrede", "Titel"
     *  - "Adresse" (+ optional zweite Spalte für PLZ/Ort etc.)
     *  - "E-Mail"/"Email"
     *  - "Halle", "Platz"
     *  - Zeitspalte (z. B. "Std-Belegung", "Zeit", "Uhrzeit") ODER getrennt "Von"/"Bis"
     *  - "Preis"/"Tarif"/"Betrag"/"Kosten"
     */
    public static CustomerData readCustomer(Path excelPath, String wantedCustomer) throws IOException {
        Objects.requireNonNull(excelPath, "excelPath must not be null");
        Objects.requireNonNull(wantedCustomer, "wantedCustomer must not be null");

        String wantedNorm = wantedCustomer.trim().toLowerCase(Locale.ROOT);

        List<BookingEntry> bookings = new ArrayList<>();

        String salutation = null; // Anrede
        String title = null;      // Titel
        String firstName = null;  // Vorname
        String lastName = null;   // Name/Nachname (aus Spalte "Name")
        String email = null;
        String address = null;

        try (InputStream in = Files.newInputStream(excelPath);
             Workbook wb = WorkbookFactory.create(in)) {

            for (String sheetName : DEFAULT_SHEETS) {
                Sheet sheet = wb.getSheet(sheetName);
                if (sheet == null) continue;

                Row header = findFirstNonEmptyRow(sheet);
                if (header == null) continue;

                // Spaltenindizes flexibel anhand Headernamen finden
                int colName   = findColumnIndex(header, List.of("name"));
                if (colName < 0) continue; // ohne Name keine Zuordnung

                int colVor    = findColumnIndex(header, List.of("vorname"));
                int colAnrede = findColumnIndex(header, List.of("anrede"));
                int colTitel  = findColumnIndex(header, List.of("titel"));

                int colAdr1   = findColumnIndex(header, List.of("adresse", "anschrift", "strasse", "straße", "addr"));
                int colAdr2   = 11; // optional

                int colMail   = findColumnIndex(header, List.of("e-mail", "email", "mail"));

                int colHalle  = findColumnIndex(header, List.of("halle"));
                int colPlatz  = findColumnIndex(header, List.of("platz", "court"));
                int colZeit   = findColumnIndex(header, List.of(
                        "std-belegung", "std.-belegung", "std . belegung", "belegung",
                        "zeit", "uhrzeit", "beginn", "start", "von"
                ));
                int colVon    = findColumnIndex(header, List.of("von", "start", "beginn", "startzeit"));
                int colBis    = findColumnIndex(header, List.of("bis", "ende", "end", "endzeit"));
                int colPreis  = findColumnIndex(header, List.of("preis", "tarif", "betrag", "kosten"));

                for (int r = header.getRowNum() + 1; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;

                    String rowName = getCellString(row.getCell(colName));
                    if (isBlank(rowName)) continue;

                    String rowNameNorm = rowName.trim().toLowerCase(Locale.ROOT);
                    if (!rowNameNorm.equals(wantedNorm)) continue; // nur gewünschter Kunde

                    // --- Stammdaten nur einmal übernehmen (erste gefundene Werte gewinnen) ---
                    if (salutation == null && colAnrede >= 0) {
                        salutation = trimOrNull(getCellString(row.getCell(colAnrede)));
                        if ("Herrn".equalsIgnoreCase(salutation)) {
                            salutation = "Herr";
                        }
                    }
                    if (title == null && colTitel >= 0)
                        title = trimOrNull(getCellString(row.getCell(colTitel)));
                    if (firstName == null && colVor >= 0)
                        firstName = trimOrNull(getCellString(row.getCell(colVor)));
                    if (lastName == null)
                        lastName = trimOrNull(rowName);

                    if (email == null && colMail >= 0) {
                        String v = trimOrNull(getCellString(row.getCell(colMail)));
                        if (!isBlank(v)) email = v;
                    }
                    if (address == null) {
                        String a1 = colAdr1 >= 0 ? trimOrNull(getCellString(row.getCell(colAdr1))) : null;
                        String a2 = colAdr2 >= 0 ? trimOrNull(getCellString(row.getCell(colAdr2))) : null;
                        String joined = joinNonBlank(" ", a1, a2);
                        if (!isBlank(joined)) address = joined;
                    }

                    // --- Buchungsfelder ---
                    String hall  = colHalle >= 0 ? trimOrNull(getCellString(row.getCell(colHalle))) : null;
                    String court = colPlatz >= 0 ? trimOrNull(getCellString(row.getCell(colPlatz))) : null;

                    // Zeit aus Zelle holen (OHNE direkt zu normalisieren)
                    String timeCellRaw = colZeit >= 0 ? trimOrNull(getCellString(row.getCell(colZeit))) : null;

                    // Erstmal roh lassen:
                    String timeRaw = null;
                    String from    = colVon  >= 0 ? trimOrNull(normalizeTimeString(getCellString(row.getCell(colVon)))) : null;
                    String to      = colBis  >= 0 ? trimOrNull(normalizeTimeString(getCellString(row.getCell(colBis)))) : null;

                    // Falls es nur eine "Zeit"-Spalte gibt:
                    if ((isBlank(from) && isBlank(to)) && !isBlank(timeCellRaw)) {
                        // Enthält eine Range? (z. B. "12:00 - 14:00",)
                        if (timeCellRaw.matches(".*\\d{1,2}:\\d{2}\\s*[-–—]\\s*\\d{1,2}:\\d{2}.*")) {
                            String[] ft = splitTimeRange(timeCellRaw);   // ["12:00","14:00"]
                            from = trimOrNull(normalizeTimeString(ft[0]));
                            to   = trimOrNull(normalizeTimeString(ft[1]));
                            // Range im Raw beibehalten (normalisiert)
                            timeRaw = from + " - " + to;
                        } else {
                            // nur ein Zeitpunkt → normalisieren und als raw übernehmen
                            String one = normalizeTimeString(timeCellRaw);
                            timeRaw = one;
                            // from/to bleiben ggf. null
                        }
                    } else {
                        // Wir hatten separate Von/Bis-Spalten -> timeRaw wenn möglich aus from/to bauen
                        if (!isBlank(from) && !isBlank(to)) {
                            timeRaw = from + " - " + to;
                        } else if (!isBlank(from)) {
                            timeRaw = from;
                        } else if (!isBlank(to)) {
                            timeRaw = to;
                        }
                    }

                    // Falls es nur eine "Zeit"-Spalte mit "12:00 - 14:00" gibt, splitten
                    if ((isBlank(from) && isBlank(to)) && !isBlank(timeRaw)) {
                        if (timeRaw.matches(".*\\d{1,2}:\\d{2}\\s*-\\s*\\d{1,2}:\\d{2}.*")) {
                            String[] ft = splitTimeRange(timeRaw);
                            from = ft[0];
                            to   = ft[1];
                        }
                    }

                    Double price = null;
                    if (colPreis >= 0) {
                        String p = getCellString(row.getCell(colPreis));
                        price = parsePriceSafe(p);
                    }

                    bookings.add(new BookingEntry(
                            sheetName,
                            r + 1,
                            hall,
                            court,
                            from,
                            to,
                            timeRaw,
                            price
                    ));
                }
            }
        }

        // Fallbacks, falls nicht gefunden
        if (lastName == null) lastName = wantedCustomer;

        return new CustomerData(salutation, title, firstName, lastName, email, address, bookings);
    }

    // ------------------------------------------------------------
    // Helpers
    // ------------------------------------------------------------

    static Row findFirstNonEmptyRow(Sheet sheet) {
        for (Row row : sheet) {
            if (row == null) continue;
            if (!isRowEmpty(row)) return row;
        }
        return null;
    }

    static boolean isRowEmpty(Row row) {
        if (row == null) return true;
        short first = row.getFirstCellNum();
        short last = row.getLastCellNum();
        if (first < 0 || last < 0) return true;

        for (int c = first; c < last; c++) {
            Cell cell = row.getCell(c);
            if (cell != null && cell.getCellType() != CellType.BLANK && !getCellString(cell).isBlank()) {
                return false;
            }
        }
        return true;
    }

    /** Sucht eine der möglichen Headerbezeichnungen (case-insensitive). */
    static int findColumnIndex(Row headerRow, List<String> wantedHeaders) {
        short first = headerRow.getFirstCellNum();
        short last = headerRow.getLastCellNum();
        if (first < 0 || last < 0) return -1;

        Set<String> targets = new HashSet<>();
        for (String h : wantedHeaders) targets.add(h.trim().toLowerCase(Locale.ROOT));

        for (int c = first; c < last; c++) {
            Cell cell = headerRow.getCell(c);
            if (cell == null) continue;
            String val = getCellString(cell);
            if (val == null) continue;
            String norm = val.trim().toLowerCase(Locale.ROOT);
            if (targets.contains(norm)) return c;
        }
        return -1;
    }

    static String getCellString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    // Excel-Datetime → ISO "yyyy-MM-ddTHH:mm:ss" (wir schneiden später auf HH:mm)
                    yield cell.getLocalDateTimeCellValue().toString();
                } else {
                    double d = cell.getNumericCellValue();
                    if (Math.floor(d) == d) {
                        yield String.valueOf((long) d);
                    } else {
                        yield String.valueOf(d);
                    }
                }
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> {
                try {
                    FormulaEvaluator evaluator = cell.getSheet().getWorkbook()
                            .getCreationHelper().createFormulaEvaluator();
                    CellValue cv = evaluator.evaluate(cell);
                    yield cv == null ? "" : switch (cv.getCellType()) {
                        case STRING -> cv.getStringValue();
                        case NUMERIC -> {
                            if (DateUtil.isCellDateFormatted(cell)) {
                                yield cell.getLocalDateTimeCellValue().toString();
                            } else {
                                yield String.valueOf(cv.getNumberValue());
                            }
                        }
                        case BOOLEAN -> String.valueOf(cv.getBooleanValue());
                        default -> "";
                    };
                } catch (Exception e) {
                    yield cell.getCellFormula();
                }
            }
            default -> "";
        };
    }

    static boolean isBlank(String s) {
        return s == null || s.trim().isEmpty();
    }
    static String trimOrNull(String s) {
        return isBlank(s) ? null : s.trim();
    }

    /** Akzeptiert u. a.:
     *  - "HH:mm", "HH:mm:ss"
     *  - "yyyy-MM-ddTHH:mm" oder "yyyy-MM-ddTHH:mm:ss" (kommt von getCellString bei Excel-Zeitformaten)
     *  - trimmt und kürzt immer auf "HH:mm"
     */
    static String normalizeTimeString(String s) {
        if (isBlank(s)) return null;
        String t = s.trim();

        // ISO-ähnlich: 2025-01-01T12:00 oder 1899-12-31T07:00:00
        if (t.matches("^\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}(:\\d{2})?$")) {
            return t.substring(11, 16);
        }

        // "HH:mm:ss" -> "HH:mm"
        if (t.matches("^\\d{1,2}:\\d{2}:\\d{2}$")) {
            return t.substring(0, 5);
        }

        // "HH:mm" -> "HH:mm"
        if (t.matches("^\\d{1,2}:\\d{2}$")) {
            return t;
        }

        // Fallback: Entferne alles außer Ziffern und Doppelpunkte, versuche erneut
        String onlyTime = t.replaceAll("[^0-9:]", "");
        if (onlyTime.matches("^\\d{1,2}:\\d{2}(:\\d{2})?$")) {
            return onlyTime.length() > 5 ? onlyTime.substring(0,5) : onlyTime;
        }

        return null;
    }

    /** Parst "12:00 - 14:00" in [from, to]; gibt nulls bei Problemen zurück. */
    static String[] splitTimeRange(String timeRaw) {
        if (isBlank(timeRaw)) return new String[]{null, null};
        String norm = timeRaw.replace('\u2013', '-')  // En dash → minus
                             .replace('\u2014', '-')  // Em dash → minus
                             .trim();
        String[] parts = norm.split("\\s*-\\s*");
        String from = parts.length > 0 ? parts[0].trim() : null;
        String to   = parts.length > 1 ? parts[1].trim() : null;
        if (from != null && from.isEmpty()) from = null;
        if (to   != null && to.isEmpty())   to   = null;
        return new String[]{from, to};
    }

    /** Parst "290", "290,00", "290.00", "290 €" -> Double; bei Fehlern null. */
    static Double parsePriceSafe(String raw) {
        if (isBlank(raw)) return null;
        String s = raw.trim()
                .replace("€", "")
                .replace("EUR", "")
                .replaceAll("\\s+", "")
                .replace(".", "")      // Tausenderpunkt entfernen
                .replace(",", ".");    // deutsches Komma in Punkt
        try {
            return Double.parseDouble(s);
        } catch (NumberFormatException e) {
            return null;
        }
    }

    // nützlich für Zusammenführung von Adresse1 + Adresse2
    static String joinNonBlank(String sep, String... parts) {
        StringBuilder sb = new StringBuilder();
        for (String p : parts) {
            if (p == null || p.isBlank()) continue;
            if (sb.length() > 0) sb.append(sep);
            sb.append(p.trim());
        }
        return sb.length() == 0 ? null : sb.toString();
    }
}
