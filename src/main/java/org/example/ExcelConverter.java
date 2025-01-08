package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelConverter {
    public static void main(String[] args) {
        SwingUtilities.invokeLater(ExcelConverter::createAndShowGUI);
    }

    public static void createAndShowGUI() {
        JFrame frame = new JFrame("Reysas ");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 600);

        String iconPath = "src/main/resources/icon.png";
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


            String orderDate = model.getValueAt(1, 0).toString(); // A2 is row 1, column 0


            String a1Value = model.getValueAt(0, 0).toString();
            String cargoNo = extractCargoNo(a1Value);


            for (int i = 1; i < model.getRowCount(); i++) {
                Row row = sheet.createRow(i);

                row.createCell(0).setCellValue("Toyota");
                row.createCell(1).setCellValue("00005");
                row.createCell(2).setCellValue("Oluşturuldu");
                row.createCell(3).setCellValue("Müşteriden Alınacak");
                row.createCell(4).setCellValue("Parsiyel");
                row.createCell(5).setCellValue(orderDate);
                row.createCell(6).setCellValue("0005");

                String dealerName = model.getValueAt(i, 3).toString();
                String[] dealerNameParts = dealerName.split(" ");
                row.createCell(7).setCellValue(dealerNameParts.length > 0 ? dealerNameParts[0] : "");
                row.createCell(8).setCellValue(dealerNameParts.length > 1 ? dealerNameParts[1] : "");
                row.createCell(9).setCellValue(dealerNameParts.length > 0 ? dealerNameParts[0] : "");
                row.createCell(10).setCellValue(model.getValueAt(i, 6).toString());
                row.createCell(11).setCellValue("");
                row.createCell(12).setCellValue("");
                row.createCell(13).setCellValue(cargoNo);
                row.createCell(14).setCellValue(model.getValueAt(i, 10).toString());
                row.createCell(15).setCellValue(model.getValueAt(i, 8).toString());
                row.createCell(16).setCellValue(dealerNameParts.length > 0 ? dealerNameParts[0] : "");
                row.createCell(17).setCellValue("TOYOTA");
                row.createCell(18).setCellValue("Araç");

                String amountStr = model.getValueAt(i, 5).toString(); // F12 is row 11, column 5
                String formattedAmount = removeTrailingZeros(amountStr);


                row.createCell(19).setCellValue(formattedAmount); // Column 19 is where "Adet" is

            }


            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }

            JOptionPane.showMessageDialog(null, "File exported successfully!", "Success", JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, "Error exporting file: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }


    private static String removeTrailingZeros(String value) {

        if (value.contains(".")) {

            value = value.replaceAll("0*$", "").replaceAll("\\.$", "");
        }
        return value;
    }




    private static String extractCargoNo(String input) {

        Pattern pattern = Pattern.compile("TIR\\d+");
        Matcher matcher = pattern.matcher(input);
        if (matcher.find()) {
            return matcher.group();
        }
        return "";
    }

}
