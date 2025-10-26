package com.thk.rechner;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.datatransfer.StringSelection;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Map;

public class InvoiceAppUI extends JFrame {

    // --- UI: Datei + Kundenwahl ---
    private final JTextField pathField = new JTextField();
    private final JButton browseBtn = new JButton("Excel wählen…");
    private final JButton reloadBtn = new JButton("Neu laden");
    private final JComboBox<String> customerCombo = new JComboBox<>();
    private final JButton showInfoBtn = new JButton("Kundeninfo anzeigen");
    private final JButton createInvoiceBtn = new JButton("Rechnung erstellen");

    // --- UI: Kundenkopf + Tabelle ---
    private final JLabel lblAnrede = new JLabel("-");
    private final JLabel lblTitel = new JLabel("-");
    private final JLabel lblVorname = new JLabel("-");
    private final JLabel lblNachname = new JLabel("-");
    private final JLabel lblAdresse = new JLabel("-");
    private final JLabel lblEmail = new JLabel("-");
    private final JButton copyEmailBtn = new JButton("📋"); // E-Mail in Zwischenablage
    

    private final DefaultTableModel bookingModel = new DefaultTableModel(
            new Object[]{"Spieltag", "Halle", "Platz", "Zeit", "Preis/Einheit"}, 0
    ) {
        @Override
        public boolean isCellEditable(int row, int column) { return false; }
    };
    private final JTable bookingTable = new JTable(bookingModel);

    private Path currentExcelPath;

    public InvoiceAppUI(Path excelPath) {
        super("Rechnungsersteller – Übersicht");
        
        try {
            java.net.URL iconURL = getClass().getClassLoader().getResource("IconTHC.png");
            if (iconURL != null) {
                ImageIcon icon = new ImageIcon(iconURL);
                setIconImage(icon.getImage());
            } else {
                System.err.println("Icon nicht gefunden: IconTHC.png");
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        
        this.currentExcelPath = excelPath;

        // ---- Grundlayout ----
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setLayout(new BorderLayout(12, 12));
        add(buildNorthPanel(), BorderLayout.NORTH);
        add(buildCenterPanel(), BorderLayout.CENTER);
        add(buildSouthPanel(), BorderLayout.SOUTH);

        // ---- Wiring ----
        pathField.setText(excelPath != null ? excelPath.toString() : "");
        createInvoiceBtn.setEnabled(false); // erst nach Info-Anzeige sinnvoll

        browseBtn.addActionListener(e -> {
            JFileChooser fc = new JFileChooser();
            if (currentExcelPath != null) {
                fc.setSelectedFile(currentExcelPath.toFile());
                File parent = currentExcelPath.toFile().getParentFile();
                if (parent != null && parent.exists()) {
                    fc.setCurrentDirectory(parent); // Komfort: direkt in den Excel-Ordner
                }
            }
            int res = fc.showOpenDialog(this);
            if (res == JFileChooser.APPROVE_OPTION) {
                currentExcelPath = fc.getSelectedFile().toPath();
                pathField.setText(currentExcelPath.toString());
                reloadCustomers();
                clearCustomerDetails();
            }
        });

        reloadBtn.addActionListener(e -> {
            currentExcelPath = Path.of(pathField.getText().trim());
            reloadCustomers();
            clearCustomerDetails();
        });

        showInfoBtn.addActionListener(e -> {
            String customer = (String) customerCombo.getSelectedItem();
            if (customer == null || customer.isBlank()) {
                info("Bitte einen Kunden auswählen.");
                return;
            }
            if (currentExcelPath == null || !Files.exists(currentExcelPath)) {
                info("Bitte zuerst eine gültige Excel-Datei wählen.");
                return;
            }
            try {
                ExcelReader.CustomerData data = ExcelReader.readCustomer(currentExcelPath, customer);
                fillCustomerDetails(data);
                createInvoiceBtn.setEnabled(true);
            } catch (Exception ex) {
                ex.printStackTrace();
                error("Fehler beim Einlesen der Kundendaten:\n" + ex.getMessage());
            }
        });

        createInvoiceBtn.addActionListener(e -> {
            String customer = (String) customerCombo.getSelectedItem();
            if (customer == null || customer.isBlank()) {
                info("Bitte einen Kunden auswählen.");
                return;
            }
            if (currentExcelPath == null || !Files.exists(currentExcelPath)) {
                info("Bitte zuerst eine gültige Excel-Datei wählen.");
                return;
            }

            try {
                // 1) Kundendaten lesen
                ExcelReader.CustomerData data = ExcelReader.readCustomer(currentExcelPath, customer);

                // 2) Template-Pfad (fixer Ort unter Desktop\Rechnungen\WordTemplate\THK_xxx-2025_name.docx)
                Path templatePath = Paths.get(
                        System.getProperty("user.home"),
                        "Desktop", "Rechnungen", "WordTemplate", "THK_xxx-2025_name.docx"
                );
                if (!Files.exists(templatePath)) {
                    error("Word Template nicht gefunden:\n" + templatePath);
                    return;
                }

                // 3) Rechnung erzeugen
                Path outFile = InvoiceGenerator.generate(data, templatePath);

                info("Rechnung erzeugt:\n" + outFile.toString());

                // Optional: Ordner im Explorer öffnen
                try {
                    Desktop.getDesktop().open(outFile.getParent().toFile());
                } catch (Exception ignore) { /* nicht kritisch */ }

            } catch (Exception ex) {
                ex.printStackTrace();
                error("Fehler bei der Rechnungserstellung:\n" + ex.getMessage());
            }
        });

        // Kopieren-Button für E-Mail
        copyEmailBtn.setToolTipText("E-Mail in die Zwischenablage kopieren");
        copyEmailBtn.addActionListener(e -> {
            String mail = lblEmail.getText();
            if (mail != null && !mail.isBlank() && !"-".equals(mail)) {
                Toolkit.getDefaultToolkit().getSystemClipboard()
                        .setContents(new StringSelection(mail), null);
            }
        });

        pack();
        setSize(900, 560);
        setLocationRelativeTo(null);

        // Initial
        if (excelPath != null) reloadCustomers();
    }

    // ---------------- UI-Bau ----------------

    private JPanel buildNorthPanel() {
        JPanel north = new JPanel(new GridBagLayout());
        GridBagConstraints gc = new GridBagConstraints();
        gc.insets = new Insets(6, 6, 6, 6);
        gc.fill = GridBagConstraints.HORIZONTAL;

        // Zeile 1: Dateipfad + Buttons
        gc.gridx = 0; gc.gridy = 0; gc.weightx = 0;
        north.add(new JLabel("Excel-Datei:"), gc);

        gc.gridx = 1; gc.weightx = 1;
        north.add(pathField, gc);

        JPanel fileBtns = new JPanel(new FlowLayout(FlowLayout.RIGHT, 6, 0));
        fileBtns.add(browseBtn);
        fileBtns.add(reloadBtn);
        gc.gridx = 2; gc.weightx = 0;
        north.add(fileBtns, gc);

        // Zeile 2: Kundenwahl + Anzeigen
        gc.gridx = 0; gc.gridy = 1; gc.weightx = 0;
        north.add(new JLabel("Kunde:"), gc);

        gc.gridx = 1; gc.weightx = 1;
        north.add(customerCombo, gc);

        gc.gridx = 2; gc.weightx = 0;
        north.add(showInfoBtn, gc);

        return north;
    }

    private JComponent buildCenterPanel() {
        JPanel center = new JPanel(new BorderLayout(10, 10));

        // Kopfkarte mit Kundendaten
        JPanel card = new JPanel(new GridBagLayout());
        card.setBorder(BorderFactory.createTitledBorder("Kunden-Infos"));

        GridBagConstraints gc = new GridBagConstraints();
        gc.insets = new Insets(4, 8, 4, 8);
        gc.anchor = GridBagConstraints.WEST;

        int r = 0;
        addRow(card, gc, r++, "Anrede:", lblAnrede, "Titel:", lblTitel);
        addRow(card, gc, r++, "Vorname:", lblVorname, "Name:",  lblNachname);
        addRow(card, gc, r++, "Adresse:", lblAdresse, "E-Mail:", buildEmailWithCopy());

        center.add(card, BorderLayout.NORTH);

        // Tabelle der Buchungen
        bookingTable.setFillsViewportHeight(true);
        bookingTable.setRowHeight(22);
        JScrollPane sp = new JScrollPane(bookingTable);
        sp.setBorder(BorderFactory.createTitledBorder("Buchungen (nicht zusammengefasst)"));
        center.add(sp, BorderLayout.CENTER);

        return center;
    }

    private JPanel buildSouthPanel() {
        JPanel south = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 6));
        south.add(createInvoiceBtn);
        return south;
    }

    private void addRow(JPanel parent, GridBagConstraints gc, int row,
                        String l1, JComponent v1, String l2, JComponent v2) {
        gc.gridx = 0; gc.gridy = row; gc.weightx = 0;
        parent.add(new JLabel(l1), gc);
        gc.gridx = 1; gc.weightx = 1;
        parent.add(v1, gc);
        gc.gridx = 2; gc.weightx = 0;
        parent.add(new JLabel(l2), gc);
        gc.gridx = 3; gc.weightx = 1;
        parent.add(v2, gc);
    }

    private JComponent buildEmailWithCopy() {
        JPanel p = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 0));
        p.add(lblEmail);
        copyEmailBtn.setMargin(new Insets(2, 6, 2, 6));
        copyEmailBtn.setEnabled(false);
        p.add(copyEmailBtn);
        return p;
    }

    // ---------------- Datenlade-Logik ----------------

    private void reloadCustomers() {
        DefaultComboBoxModel<String> model = new DefaultComboBoxModel<>();
        customerCombo.setModel(model);

        if (currentExcelPath == null || !Files.exists(currentExcelPath)) {
            error("Excel-Datei nicht gefunden:\n" + (currentExcelPath == null ? "" : currentExcelPath));
            return;
        }

        try {
            Map<String, List<ExcelReader.RowRef>> grouped =
                    ExcelReader.groupRowsByCustomerAllSheets(currentExcelPath);

            grouped.keySet().stream()
                    .sorted(String.CASE_INSENSITIVE_ORDER)
                    .forEach(model::addElement);

            if (model.getSize() > 0) customerCombo.setSelectedIndex(0);
        } catch (Exception ex) {
            ex.printStackTrace();
            error("Fehler beim Laden der Kundennamen:\n" + ex.getMessage());
        }
    }

    private void clearCustomerDetails() {
        lblAnrede.setText("-");
        lblTitel.setText("-");
        lblVorname.setText("-");
        lblNachname.setText("-");
        lblAdresse.setText("-");
        lblEmail.setText("-");
        copyEmailBtn.setEnabled(false);
        bookingModel.setRowCount(0);
        createInvoiceBtn.setEnabled(false);
    }

    private void fillCustomerDetails(ExcelReader.CustomerData data) {
        // Kopf
        lblAnrede.setText(nvl(data.salutation()));
        lblTitel.setText(nvl(data.title()));
        lblVorname.setText(nvl(data.firstName()));
        lblNachname.setText(nvl(data.lastName()));
        lblAdresse.setText(nvl(data.address()));
        lblEmail.setText(nvl(data.email()));
        copyEmailBtn.setEnabled(data.email() != null && !data.email().isBlank());

        // Tabelle
        bookingModel.setRowCount(0);
        for (ExcelReader.BookingEntry b : data.bookings()) {
            String zeit = buildTimeRange(b.timeFrom(), b.timeTo(), b.timeRaw());
            bookingModel.addRow(new Object[]{
                    b.sheet(),
                    nvl(b.hall()),
                    nvl(b.court()),
                    zeit,
                    formatPrice(b.price())
            });
        }
    }

    // ---------------- Hilfsfunktionen ----------------

    private static String nvl(String v) { return (v == null || v.isBlank()) ? "-" : v; }

    private static String formatPrice(Double p) {
        if (p == null) return "-";
        return String.format("%.2f €", p);
    }

    private static String buildTimeRange(String from, String to, String raw) {
        // bevorzugt from-to, sonst raw, sonst "-"
        if (from != null && !from.isBlank() && to != null && !to.isBlank()) {
            return from + " - " + to;
        }
        if (raw != null && !raw.isBlank()) return raw;
        return "-";
    }

    // ---------------- main (zum direkten Starten) ----------------

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            Path excel = null;

            // 1) Falls per Argument übergeben und vorhanden -> nutzen
            if (args != null && args.length > 0) {
                Path candidate = Path.of(args[0]);
                if (Files.exists(candidate)) {
                    excel = candidate;
                }
            }

            // 2) Sonst: dynamischer Benutzerpfad (USERPROFILE) -> Desktop\Rechnungen\Excel\THK_...
            if (excel == null) {
                Path candidate = Paths.get(
                        System.getProperty("user.home"),
                        "Desktop", "Rechnungen", "Excel", "THK_Belegungsplan_ReAus_Programm.xlsx"
                );
                if (Files.exists(candidate)) {
                    excel = candidate;
                }
            }

            // 3) Fallback: lokale Datei im Arbeitsverzeichnis (wie bisher)
            if (excel == null) {
                Path candidate = Path.of("THK_Belegungsplan_ReAus_Programm.xlsx");
                if (Files.exists(candidate)) {
                    excel = candidate;
                }
            }

            new InvoiceAppUI(excel).setVisible(true);
        });
    }

    // --- kleine Dialoghelfer ---
    private void info(String msg) {
        JOptionPane.showMessageDialog(this, msg, "Hinweis", JOptionPane.INFORMATION_MESSAGE);
    }
    private void error(String msg) {
        JOptionPane.showMessageDialog(this, msg, "Fehler", JOptionPane.ERROR_MESSAGE);
    }
}
