package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.apache.commons.lang3.StringUtils.defaultString;

public class ExcelConverter {
    public static void main(String[] args) {
        SwingUtilities.invokeLater(ExcelConverter::createAndShowGUI);
    }

    public static void createAndShowGUI() {
        JFrame frame = new JFrame("Reysas ");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(900, 700);

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
            Image scaledImage = logoIcon.getImage().getScaledInstance(500, 150, Image.SCALE_SMOOTH);
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
            fileChooser.setFileFilter(new javax.swing.filechooser.FileFilter() {
                @Override
                public boolean accept(File file) {
                    // Allow directories and .xls/.xlsx files
                    return file.isDirectory() || file.getName().toLowerCase().endsWith(".xls") || file.getName().toLowerCase().endsWith(".xlsx");
                }

                @Override
                public String getDescription() {
                    return "Excel Files (*.xls, *.xlsx)";
                }
            });
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
            if (!isValidDate(sheet.getRow(1))) {
                JOptionPane.showMessageDialog(null, "Yanlış Dosya Formatı!", "Hata", JOptionPane.ERROR_MESSAGE);
                return;
            }


            int columnCount = sheet.getRow(0).getPhysicalNumberOfCells();
            for (int i = 0; i < columnCount; i++) {
                char columnLetter = (char) ('A' + i);
                model.addColumn(String.valueOf(columnLetter));
            }


            for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) continue;

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

    private static boolean isValidDate(Row row) {
        if (row == null) return false;


        Cell cell = row.getCell(0);


        return isDate(cell);
    }

    private static boolean isDate(Cell cell) {
        if (cell == null) return false;
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            return true;
        }
        if (cell.getCellType() == CellType.STRING) {
            String cellValue = cell.getStringCellValue();
            try {
                SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
                sdf.setLenient(false);
                sdf.parse(cellValue);
                return true;
            } catch (ParseException e) {
                return false;
            }
        }
        return false;
    }


    private static Object getCellValue(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return characterConverter(cell.getStringCellValue());
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

    private static String characterConverter(String text) {
        if (text == null) return "";
        return text
                .replace("Ð", "Ğ")
                .replace("ð", "ğ")
                .replace("ý", "ı")
                .replace("Ý", "İ")
                .replace("þ", "ş")
                .replace("Þ", "Ş")
                .replace("Û", "Ü")
                .replace("û", "ü")
                .replace("Ò", "Ö")
                .replace("ò", "ö")
                .replace("Ýçel", "Mersin")
                .replace("İçel", "Mersin");

    }

    public static void convertToOutputFormat(JTable table, File outputFile) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Kitap1");

            String[] headers = {
                    "Proje", "Müşteri", "Sipariş Durumu", "Sipariş Türü", "Yükleme Tipi",
                    "Sipariş Tarihi", "Yükleme Firması", "Yükleme Firması Adres Tipi",
                    "Boşaltma Firması", "Boşaltma Firması Adres Tipi",
                    "Müşteri İrsaliye", "İrsaliye seri", "İrsaliye no", "Yük Numarası",
                    "Model", "Şasi No", "Lokasyon", "Marka", "Kap Cinsi", "Adet"
            };

            DefaultTableModel model = (DefaultTableModel) table.getModel();
            int rowIndex = 0;
            String previousInvoiceNo = null;
            boolean firstDataRow = true;


            for (int i = 1; i < model.getRowCount(); i++) {


                String materialName         = getValue(model, i, 10);
                String colorCode       = getValue(model, i, 9);
                String currentInvoice = getValue(model, i, 6);
                String firstCell      = getValue(model, i, 0);
                //String dealerName     = getValue(model, i, 3);
                String date = getValue(model,0,0);
                String[] dateParts = date.split("\\s+");
                String dateCellFirstPart = (dateParts.length > 0) ? dateParts[0] : "";
                if (!dateCellFirstPart.isEmpty()) {
                    dateCellFirstPart = dateCellFirstPart.replace("/", ".");
                }

                if (firstCell.trim().equalsIgnoreCase("Proje")) {
                    continue;
                }


                if (materialName.toUpperCase().contains("MAL AD")) {
                    continue;
                }


                if (colorCode.toUpperCase().contains("RENK KODU")) {
                    continue;
                }


                if (materialName.isEmpty() && colorCode.isEmpty()) {
                    continue;
                }


                String amount = getValue(model, i, 6);
                if (amount.isEmpty()) {
                    continue;
                }



                if (firstDataRow) {

                    Row headerRow = sheet.createRow(rowIndex++);
                    for (int j = 0; j < headers.length; j++) {
                        headerRow.createCell(j).setCellValue(headers[j]);
                    }
                    firstDataRow = false;
                } else {

                    if (!currentInvoice.equals(previousInvoiceNo))
                    {
                        sheet.createRow(rowIndex++);
                        Row headerRow2 = sheet.createRow(rowIndex++);
                        for (int j = 0; j < headers.length; j++) {
                            headerRow2.createCell(j).setCellValue(headers[j]);
                        }
                    }
                }


                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue("Toyota");
                row.createCell(1).setCellValue("00005");
                row.createCell(2).setCellValue("Oluşturuldu");
                row.createCell(3).setCellValue("Müşteriden Teslim Alınacak");
                row.createCell(4).setCellValue("Parsiyel");


                row.createCell(5).setCellValue(dateCellFirstPart);

                row.createCell(6).setCellValue("0005");


                //String[] dealerNameParts = dealerName.split(" ");

                row.createCell(7).setCellValue(getValue(model,i,15));
                row.createCell(8).setCellValue(getValue(model,i,3));
                row.createCell(9).setCellValue(getValue(model,i,15));

                row.createCell(10).setCellValue(currentInvoice);
                row.createCell(11).setCellValue("");
                row.createCell(12).setCellValue("");


                String cargoNo = extractCargoNo(getValue(model, 0, 0));
                row.createCell(13).setCellValue(cargoNo);

                row.createCell(14).setCellValue(materialName);
                row.createCell(15).setCellValue(getValue(model, i, 8));
                /*row.createCell(16).setCellValue(
                        dealerNameParts.length > 0 ? dealerNameParts[0] : ""
                );*/
                row.createCell(16).setCellValue(getValue(model,i,3));
                row.createCell(17).setCellValue("TOYOTA");
                row.createCell(18).setCellValue("Araç");
                row.createCell(19).setCellValue("1");

                previousInvoiceNo = currentInvoice;
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }

            JOptionPane.showMessageDialog(null,
                    "Dosya başarıyla kaydedildi!",
                    "Başarılı",
                    JOptionPane.INFORMATION_MESSAGE);

        } catch (IOException e) {
            JOptionPane.showMessageDialog(null,
                    "Dosya kaydedilemedi!: " + e.getMessage(),
                    "Hata:",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private static String getValue(DefaultTableModel model, int row, int col) {
        if (row < 0 || row >= model.getRowCount()) return "";
        if (col < 0 || col >= model.getColumnCount()) return "";
        Object val = model.getValueAt(row, col);
        return characterConverter((val == null) ? "" : val.toString().trim());
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
