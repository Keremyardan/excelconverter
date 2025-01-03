package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.*;

public class ExcelConverter {
    public static void main(String[] args) {
        SwingUtilities.invokeLater(ExcelConverter::createAndShowGUI);
    }

    public static void createAndShowGUI() {
        JFrame frame = new JFrame("Reysas ");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 600);

        String iconPath = "src/main/resources/icon.png"; // Adjust path if needed
        File iconFile = new File(iconPath);
        if (iconFile.exists()) {
            ImageIcon icon = new ImageIcon(iconFile.getAbsolutePath());
            frame.setIconImage(icon.getImage());
        }

        frame.setLocationRelativeTo(null);

        JPanel panel = new JPanel(new BorderLayout());

        String logoPath = "src/main/resources/logo.png";
        JLabel logoLabel = new JLabel();
        logoLabel.setHorizontalAlignment(SwingConstants.CENTER);
        File logoFile = new File(logoPath);

        if (logoFile.exists()) {
            ImageIcon logoIcon = new ImageIcon(logoFile.getAbsolutePath());
            Image scaledImage = logoIcon.getImage().getScaledInstance(400, 100, Image.SCALE_SMOOTH);
            logoLabel.setIcon(new ImageIcon(scaledImage));
        } else {
            logoLabel.setText("Logo not found!");
        }

        panel.add(logoLabel, BorderLayout.NORTH);

        JButton openButton = new JButton("Excel Dosyasını Aç");
        JButton exportButton = new JButton("Dosyayı Kaydet");

        exportButton.setEnabled(false);
        JTable table = new JTable();
        JScrollPane scrollPane = new JScrollPane(table);

        openButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            int option = fileChooser.showOpenDialog(frame);
            if (option == JFileChooser.APPROVE_OPTION) {
                File file = fileChooser.getSelectedFile();
                loadExcelToTable(file, table);
                exportButton.setEnabled(true);
            }
        });

        exportButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            int option = fileChooser.showSaveDialog(frame);
            if (option == JFileChooser.APPROVE_OPTION) {
                File file = fileChooser.getSelectedFile();
                if (!file.getName().toLowerCase().endsWith(".xlsx")) {
                    file = new File(file.getAbsolutePath() + ".xlsx");
                }
                convertToOutputFormat(table, file);
            }
        });

        JPanel buttonPanel = new JPanel();
        buttonPanel.add(openButton);
        buttonPanel.add(exportButton);

        panel.add(buttonPanel, BorderLayout.SOUTH);
        panel.add(scrollPane, BorderLayout.CENTER);

        frame.add(panel);
        frame.setVisible(true);
    }

    public static void loadExcelToTable(File file, JTable table) {
        try (FileInputStream fis = new FileInputStream(file); Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            DefaultTableModel model = new DefaultTableModel();

            int columnCount = sheet.getRow(0).getPhysicalNumberOfCells();
            for (int i = 0; i < columnCount; i++) {
                char columnLetter = (char) ('A' + i);
                model.addColumn(String.valueOf(columnLetter));
            }

            for (Row row : sheet) {
                int cellCount = row.getLastCellNum();
                Object[] rowData = new Object[cellCount];
                for (int i = 0; i < cellCount; i++) {
                    Cell cell = row.getCell(i);
                    rowData[i] = getCellValue(cell);
                }
                model.addRow(rowData);
            }

            table.setModel(model);
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, "Error loading file: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private static Object getCellValue(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return "";
        }
    }

    public static void convertToOutputFormat(JTable table, File outputFile) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Converted Data");

            String[] headers = {
                    "Proje", "Müşteri", "Sipariş Durumu", "Sipariş Türü", "Yükleme Tipi",
                    "Sipariş Tarihi", "Yükleme Firması", "Yükleme Firması Adres Tipi",
                    "Boşaltma Firması", "Boşaltma Firması Adres Tipi", "Müşteri İrsaliye",
                    "İrsaliye seri", "İrsaliye no", "Yük Numarası", "Model", "Şasi No",
                    "Lokasyon", "Marka", "Kap Cinsi", "Adet"
            };

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                headerRow.createCell(i).setCellValue(headers[i]);
            }

            DefaultTableModel model = (DefaultTableModel) table.getModel();

            // Get the date from A2 (second row, first column)
            String siparisTarihi = model.getValueAt(1, 0).toString(); // A2 is row 1, column 0

            // Iterate through the rows of the input table
            for (int i = 0; i < model.getRowCount(); i++) {
                Row row = sheet.createRow(i + 1);

                // Proje: Always "Toyota"
                row.createCell(0).setCellValue("Toyota");
                // Müşteri: Always "00005"
                row.createCell(1).setCellValue("00005");
                // Sipariş Durumu: Always "Oluşturuldu"
                row.createCell(2).setCellValue("Oluşturuldu");
                // Sipariş Türü: Always "Müşteriden Alınacak"
                row.createCell(3).setCellValue("Müşteriden Alınacak");
                // Yükleme Tipi: Always "Parsiyel"
                row.createCell(4).setCellValue("Parsiyel");

                // Sipariş Tarihi: From A2 (siparisTarihi)
                row.createCell(5).setCellValue(siparisTarihi);

                // Yükleme Firması: Always "00005"
                row.createCell(6).setCellValue("00005");

                // Yükleme Firması Adres Tipi: First word of "Bayi Adı" (cell 1)
                String bayiAdı = model.getValueAt(4, 3).toString();
                String[] bayiAdıParts = bayiAdı.split(" ");
                row.createCell(7).setCellValue(bayiAdıParts.length > 0 ? bayiAdıParts[0] : "");

                // Boşaltma Firması: Second word of "Bayi Adı"
                row.createCell(8).setCellValue(bayiAdıParts.length > 1 ? bayiAdıParts[1] : "");

// Boşaltma Firması Adres Tipi: First word of "Bayi Adı"
                row.createCell(9).setCellValue(bayiAdıParts.length > 0 ? bayiAdıParts[0] : "");


                // Müşteri İrsaliye: From "İrsaliye No" (cell 2)
                row.createCell(10).setCellValue(model.getValueAt(i, 2).toString());

                // İrsaliye Seri: Left empty
                row.createCell(11).setCellValue("");

                // İrsaliye No: Left empty
                row.createCell(12).setCellValue("");

                // Yük Numarası: From "Tır No" (first column)
                row.createCell(13).setCellValue(model.getValueAt(i, 0).toString());

                // Model: From "Mal Adı" (cell 6)
                row.createCell(14).setCellValue(model.getValueAt(i, 6).toString());

                // Şasi No: From "Şasi No" (cell 4)
                row.createCell(15).setCellValue(model.getValueAt(i, 4).toString());

                // Lokasyon: First word of "Bayi Adı" (cell 1)
                row.createCell(16).setCellValue(bayiAdıParts.length > 0 ? bayiAdıParts[0] : "");

                // Marka: Always "TOYOTA"
                row.createCell(17).setCellValue("TOYOTA");

                // Kap Cinsi: Always "Araç"
                row.createCell(18).setCellValue("Araç");

                // Adet: From "Tırdaki Araç Sayısı" (last column, cell 19)
                int adet = 0; // You can get this value based on your specific structure
                row.createCell(19).setCellValue(adet);
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }

            JOptionPane.showMessageDialog(null, "File exported successfully!", "Success", JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, "Error exporting file: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }
}
