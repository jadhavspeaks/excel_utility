package com.excelcomparator;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelComparator extends JFrame {

    private JTable preview1;
    private JTable preview2;
    private JTable resultsTable;
    private ExcelUtil.ExcelData data1;
    private ExcelUtil.ExcelData data2;

    private JPanel columnMappingPanel;
    private List<JComboBox<String>> file1ColumnDropdowns;
    private List<JComboBox<String>> file2ColumnDropdowns;

    private JCheckBox columnLevelComparison;
    private JCheckBox detectMissingExtraRows;

    private ComparisonResult lastResult;


    public ExcelComparator() {
        setTitle("Excel Comparator");
        setSize(1280, 800);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);

        file1ColumnDropdowns = new ArrayList<>();
        file2ColumnDropdowns = new ArrayList<>();

        // Main panel
        JPanel mainPanel = new JPanel(new BorderLayout());

        // Top panel for file choosers
        JPanel topPanel = new JPanel(new GridLayout(1, 2));
        // File 1
        JPanel file1Panel = new JPanel();
        file1Panel.setBorder(BorderFactory.createTitledBorder("File 1"));
        JButton file1Button = new JButton("Choose File 1");
        file1Panel.add(file1Button);
        topPanel.add(file1Panel);

        // File 2
        JPanel file2Panel = new JPanel();
        file2Panel.setBorder(BorderFactory.createTitledBorder("File 2"));
        JButton file2Button = new JButton("Choose File 2");
        file2Panel.add(file2Button);
        topPanel.add(file2Panel);

        mainPanel.add(topPanel, BorderLayout.NORTH);

        // Center panel for previews and options
        JPanel centerPanel = new JPanel(new BorderLayout());

        // Preview tables
        JPanel previewPanel = new JPanel(new GridLayout(1, 2));
        preview1 = new JTable();
        previewPanel.add(new JScrollPane(preview1));
        preview2 = new JTable();
        previewPanel.add(new JScrollPane(preview2));
        centerPanel.add(previewPanel, BorderLayout.CENTER);

        // Options Panel
        JPanel optionsPanel = new JPanel(new BorderLayout());
        optionsPanel.setBorder(BorderFactory.createTitledBorder("Comparison Options"));

        // Column Mappings UI
        columnMappingPanel = new JPanel();
        columnMappingPanel.setLayout(new BoxLayout(columnMappingPanel, BoxLayout.Y_AXIS));
        JScrollPane mappingScrollPane = new JScrollPane(columnMappingPanel);
        mappingScrollPane.setPreferredSize(new Dimension(400, 100));

        JButton addMappingButton = new JButton("+ Add Key");
        addMappingButton.addActionListener(e -> addMappingRow());

        JPanel mappingContainer = new JPanel(new BorderLayout());
        mappingContainer.add(new JLabel("Map Key Columns:", SwingConstants.CENTER), BorderLayout.NORTH);
        mappingContainer.add(mappingScrollPane, BorderLayout.CENTER);
        mappingContainer.add(addMappingButton, BorderLayout.SOUTH);

        optionsPanel.add(mappingContainer, BorderLayout.CENTER);

        // Checkboxes
        JPanel checkboxPanel = new JPanel(new GridLayout(3, 1));
        columnLevelComparison = new JCheckBox("Column-level comparison", true);
        detectMissingExtraRows = new JCheckBox("Detect missing/extra rows", true);
        checkboxPanel.add(columnLevelComparison);
        checkboxPanel.add(detectMissingExtraRows);
        optionsPanel.add(checkboxPanel, BorderLayout.WEST);

        centerPanel.add(optionsPanel, BorderLayout.SOUTH);
        mainPanel.add(centerPanel, BorderLayout.CENTER);

        // Bottom panel for results and actions
        JPanel bottomPanel = new JPanel(new BorderLayout());
        resultsTable = new JTable();
        bottomPanel.add(new JScrollPane(resultsTable), BorderLayout.CENTER);

        JPanel actionPanel = new JPanel();
        JButton runButton = new JButton("Run Comparison");
        JButton exportButton = new JButton("Export to Excel");
        actionPanel.add(runButton);
        actionPanel.add(exportButton);
        bottomPanel.add(actionPanel, BorderLayout.SOUTH);

        mainPanel.add(bottomPanel, BorderLayout.SOUTH);

        add(mainPanel);

        // Add action listeners
        file1Button.addActionListener(this::chooseFile1);
        file2Button.addActionListener(this::chooseFile2);
        runButton.addActionListener(this::runComparison);
        exportButton.addActionListener(this::exportResults);

        // Initial mapping row
        addMappingRow();
    }

    private void runComparison(ActionEvent e) {
        if (data1 == null || data2 == null) {
            JOptionPane.showMessageDialog(this, "Please load both Excel files.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        Map<String, String> keyColumns = new HashMap<>();
        for (int i = 0; i < file1ColumnDropdowns.size(); i++) {
            String col1 = (String) file1ColumnDropdowns.get(i).getSelectedItem();
            String col2 = (String) file2ColumnDropdowns.get(i).getSelectedItem();
            if (col1 != null && col2 != null) {
                keyColumns.put(col1, col2);
            }
        }

        if (keyColumns.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please map at least one key column.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        lastResult = ComparatorLogic.compare(data1, data2, keyColumns, columnLevelComparison.isSelected(), detectMissingExtraRows.isSelected());
        displayResults(lastResult);
    }

    private void exportResults(ActionEvent e) {
        if (lastResult == null) {
            JOptionPane.showMessageDialog(this, "No comparison results to export.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Save As");
        fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Workbook (*.xlsx)", "xlsx"));
        int userSelection = fileChooser.showSaveDialog(this);

        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File fileToSave = fileChooser.getSelectedFile();
            if (!fileToSave.getPath().endsWith(".xlsx")) {
                fileToSave = new File(fileToSave.getPath() + ".xlsx");
            }
            try {
                ExcelUtil.writeResultToExcel(lastResult, data1, data2, fileToSave);
                JOptionPane.showMessageDialog(this, "Results exported successfully to " + fileToSave.getName(), "Export Successful", JOptionPane.INFORMATION_MESSAGE);
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(this, "Error exporting results: " + ex.getMessage(), "Export Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void displayResults(ComparisonResult result) {
        List<String> headers = new ArrayList<>();
        headers.add("Status");
        headers.addAll(result.getHeaders());

        DefaultTableModel model = new DefaultTableModel(headers.toArray(new String[0]), 0);

        List<String> headers1 = data1.getHeaders();
        List<String> headers2 = data2.getHeaders();

        for (ComparisonResult.ResultRow row : result.getResultRows()) {
            List<Object> rowData = new ArrayList<>();
            rowData.add(row.getStatus());
            for (String header : result.getHeaders()) {
                int idx1 = headers1.indexOf(header);
                int idx2 = headers2.indexOf(header);
                Object val;

                switch (row.getStatus()) {
                    case MISSING_IN_FILE_1:
                        val = (idx2 != -1 && idx2 < row.getData2().size()) ? row.getData2().get(idx2) : "";
                        break;
                    case MISSING_IN_FILE_2:
                        val = (idx1 != -1 && row.getData1() != null && idx1 < row.getData1().size()) ? row.getData1().get(idx1) : "";
                        break;
                    case MATCH:
                        val = (idx1 != -1 && row.getData1() != null && idx1 < row.getData1().size()) ? row.getData1().get(idx1) : "";
                        break;
                    case MISMATCH:
                        Object val1 = (idx1 != -1 && row.getData1() != null && idx1 < row.getData1().size()) ? row.getData1().get(idx1) : null;
                        Object val2 = (idx2 != -1 && row.getData2() != null && idx2 < row.getData2().size()) ? row.getData2().get(idx2) : null;
                        if (!java.util.Objects.equals(val1, val2)) {
                            val = String.format("%s -> %s", val1, val2);
                        } else {
                            val = val1;
                        }
                        break;
                    default:
                        val = "";
                }
                rowData.add(val);
            }
            model.addRow(rowData.toArray());
        }

        resultsTable.setModel(model);
        resultsTable.setDefaultRenderer(Object.class, new ResultCellRenderer(result));
    }


    private void addMappingRow() {
        JPanel row = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JComboBox<String> cb1 = new JComboBox<>();
        JComboBox<String> cb2 = new JComboBox<>();
        file1ColumnDropdowns.add(cb1);
        file2ColumnDropdowns.add(cb2);

        updateColumnMappings();

        JButton removeButton = new JButton("-");
        removeButton.addActionListener(e -> {
            columnMappingPanel.remove(row);
            file1ColumnDropdowns.remove(cb1);
            file2ColumnDropdowns.remove(cb2);
            columnMappingPanel.revalidate();
            columnMappingPanel.repaint();
        });

        row.add(new JLabel("File 1 Column:"));
        row.add(cb1);
        row.add(new JLabel("  vs  "));
        row.add(new JLabel("File 2 Column:"));
        row.add(cb2);
        row.add(removeButton);
        columnMappingPanel.add(row);
        columnMappingPanel.revalidate();
        columnMappingPanel.repaint();
    }

    private void updateColumnMappings() {
        String[] headers1 = (data1 != null) ? data1.getHeaders().toArray(new String[0]) : new String[0];
        String[] headers2 = (data2 != null) ? data2.getHeaders().toArray(new String[0]) : new String[0];

        for (JComboBox<String> cb : file1ColumnDropdowns) {
            Object selected = cb.getSelectedItem();
            cb.setModel(new DefaultComboBoxModel<>(headers1));
            cb.setSelectedItem(selected);
        }
        for (JComboBox<String> cb : file2ColumnDropdowns) {
            Object selected = cb.getSelectedItem();
            cb.setModel(new DefaultComboBoxModel<>(headers2));
            cb.setSelectedItem(selected);
        }
    }

    private void chooseFile1(ActionEvent e) {
        JFileChooser fileChooser = new JFileChooser();
        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            try {
                data1 = ExcelUtil.readExcel(selectedFile, -1);
                updatePreviewTable(preview1, data1, 10);
                updateColumnMappings();
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(this, "Error reading file: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void chooseFile2(ActionEvent e) {
        JFileChooser fileChooser = new JFileChooser();
        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            try {
                data2 = ExcelUtil.readExcel(selectedFile, -1);
                updatePreviewTable(preview2, data2, 10);
                updateColumnMappings();
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(this, "Error reading file: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void updatePreviewTable(JTable table, ExcelUtil.ExcelData excelData, int numRows) {
        if (excelData == null) {
            table.setModel(new DefaultTableModel());
            return;
        }

        String[] headers = excelData.getHeaders().toArray(new String[0]);
        List<List<Object>> dataList = excelData.getData();
        int rowsToPreview = Math.min(numRows, dataList.size());

        Object[][] data = dataList.subList(0, rowsToPreview).stream()
                .map(List::toArray)
                .toArray(Object[][]::new);

        DefaultTableModel model = new DefaultTableModel(data, headers);
        table.setModel(model);
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            ExcelComparator ex = new ExcelComparator();
            ex.setVisible(true);
        });
    }
}

class ResultCellRenderer extends DefaultTableCellRenderer {
    private final ComparisonResult result;

    public ResultCellRenderer(ComparisonResult result) {
        this.result = result;
    }

    @Override
    public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
        Component c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

        if (result != null && row < result.getResultRows().size()) {
            ComparisonResult.ResultRow resultRow = result.getResultRows().get(row);
            ComparisonResult.RowStatus status = resultRow.getStatus();
            Color backgroundColor = Color.WHITE;

            switch (status) {
                case MATCH:
                    backgroundColor = Color.WHITE;
                    break;
                case MISMATCH:
                    backgroundColor = new Color(255, 255, 224); // Light yellow
                    if (column > 0) { // Not the status column
                        List<Boolean> mismatches = resultRow.getMismatches();
                        if (mismatches != null && (column - 1) < mismatches.size() && mismatches.get(column - 1)) {
                            backgroundColor = new Color(255, 218, 185); // Peach
                        }
                    }
                    break;
                case MISSING_IN_FILE_1:
                case MISSING_IN_FILE_2:
                    backgroundColor = new Color(255, 228, 225); // Misty Rose
                    break;
            }
            c.setBackground(backgroundColor);

        } else {
            c.setBackground(Color.WHITE);
        }

        if (isSelected) {
            c.setBackground(table.getSelectionBackground());
            c.setForeground(table.getSelectionForeground());
        } else {
            c.setForeground(table.getForeground());
        }

        return c;
    }
}
