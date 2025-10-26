package com.thk.rechner;

import javax.swing.*;
import java.nio.file.Files;
import java.nio.file.Path;

public class App {

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            Path excelPath = null;

            // 1) Falls ein Pfad als Argument übergeben wurde
            if (args != null && args.length > 0) {
                Path candidate = Path.of(args[0]);
                if (Files.exists(candidate)) {
                    excelPath = candidate;
                }
            }

            // 2) Fallback: Standard-Dateiname im Arbeitsverzeichnis
            if (excelPath == null) {
                Path candidate = Path.of("THK_Belegungsplan_ReAus_Programm.xlsx");
                if (Files.exists(candidate)) {
                    excelPath = candidate;
                }
            }

            // 3) Wenn immer noch nichts da ist: Dateiauswahl anbieten
            if (excelPath == null) {
                JFileChooser fc = new JFileChooser();
                fc.setDialogTitle("Bitte Excel-Datei wählen");
                int res = fc.showOpenDialog(null);
                if (res == JFileChooser.APPROVE_OPTION) {
                    excelPath = fc.getSelectedFile().toPath();
                }
            }

            // 4) UI starten (excelPath darf auch null sein – die UI zeigt dann einen Hinweis)
            new InvoiceAppUI(excelPath).setVisible(true);
        });
    }
}
