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

        JPanel panel = new JPanel(new BorderLayout());

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
                exportTableToExcel(table, file);
            }
        });

        JPanel buttonPanel = new JPanel();
        buttonPanel.add(openButton);
        buttonPanel.add(exportButton);

        panel.add(buttonPanel, BorderLayout.NORTH);
        panel.add(scrollPane, BorderLayout.CENTER);

        frame.add(panel);
        frame.setVisible(true);
    }

    public static void loadExcelToTable(File file, JTable table) {
        try (FileInputStream fis = new FileInputStream(file); Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            DefaultTableModel model = new DefaultTableModel();

            for (Row row : sheet) {
                int cellCount = row.getLastCellNum();
                if (model.getColumnCount() < cellCount) {
                    for (int i = model.getColumnCount(); i < cellCount; i++) {
                        model.addColumn("Column " + (i + 1));
                    }
                }

                Object[] rowData = new Object[cellCount];
                for (int i = 0; i < cellCount; i++) {
                    Cell cell = row.getCell(i);
                    rowData[i] = cell != null ? cell.toString() : "";
                }
                model.addRow(rowData);
            }

            table.setModel(model);
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, "Error loading file: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    public static void exportTableToExcel(JTable table, File file) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Exported Data");
            DefaultTableModel model = (DefaultTableModel) table.getModel();

            for (int i = 0; i < model.getRowCount(); i++) {
                Row row = sheet.createRow(i);
                for (int j = 0; j < model.getColumnCount(); j++) {
                    Cell cell = row.createCell(j);
                    Object value = model.getValueAt(i, j);
                    cell.setCellValue(value != null ? value.toString() : "");
                }
            }

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
            JOptionPane.showMessageDialog(null, "File exported successfully!", "Success", JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, "Error exporting file: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }
}
