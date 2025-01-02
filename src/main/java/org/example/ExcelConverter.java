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

            for (int i = 0; i < model.getRowCount(); i++) {
                Row row = sheet.createRow(i + 1);

                row.createCell(0).setCellValue("Toyota"); // Proje
                row.createCell(1).setCellValue(""); // Müşteri
                row.createCell(2).setCellValue("Oluşturuldu"); // Sipariş Durumu
                row.createCell(3).setCellValue(""); // Sipariş Türü
                row.createCell(4).setCellValue(""); // Yükleme Tipi

                row.createCell(5).setCellValue(model.getValueAt(i, 0).toString()); // Sipariş Tarihi (A2 den gelen veri)
                row.createCell(6).setCellValue(""); // Yükleme Firması
                row.createCell(7).setCellValue(""); // Yükleme Firması Adres Tipi
                row.createCell(8).setCellValue(""); // Boşaltma Firması
                row.createCell(9).setCellValue(""); // Boşaltma Firması Adres Tipi

                row.createCell(10).setCellValue(""); // Müşteri İrsaliye
                row.createCell(11).setCellValue(""); // İrsaliye seri
                row.createCell(12).setCellValue(model.getValueAt(i, 2).toString()); // İrsaliye no
                row.createCell(13).setCellValue(""); // Yük Numarası
                row.createCell(14).setCellValue(""); // Model

                row.createCell(15).setCellValue(model.getValueAt(i, 4).toString()); // Şasi No
                row.createCell(16).setCellValue(model.getValueAt(i, 5).toString()); // Lokasyon
                row.createCell(17).setCellValue("TOYOTA"); // Marka
                row.createCell(18).setCellValue("Araç"); // Kap Cinsi
                row.createCell(19).setCellValue(8); // Adet (örnekte sabit)
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
