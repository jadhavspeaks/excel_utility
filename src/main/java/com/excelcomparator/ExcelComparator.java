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
import java.util.Arrays;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class ExcelComparator extends JFrame {

    private JTable preview1, preview2, resultsTable;
    private ExcelUtil.ExcelData data1, data2;
    private File file1, file2;
    private JComboBox<String> sheet1ComboBox, sheet2ComboBox;
    private JSpinner headerRowsSpinner1, headerRowsSpinner2;

    private JCheckBox chkNormalize1, chkNormalize2;

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

        JPanel mainPanel = new JPanel(new BorderLayout(5, 5));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));

        JPanel topPanel = new JPanel(new GridLayout(1, 2, 5, 5));

        JPanel file1Panel = createFilePanel(1);
        JPanel file2Panel = createFilePanel(2);

        topPanel.add(file1Panel);
        topPanel.add(file2Panel);

        mainPanel.add(topPanel, BorderLayout.NORTH);

        JSplitPane centerSplit = new JSplitPane(JSplitPane.VERTICAL_SPLIT);
        centerSplit.setResizeWeight(0.6);
        mainPanel.add(centerSplit, BorderLayout.CENTER);

        JPanel previewPanel = new JPanel(new GridLayout(1, 2, 5, 5));
        previewPanel.setBorder(BorderFactory.createTitledBorder("Preview (First 10 Rows)"));
        preview1 = new JTable();
        previewPanel.add(new JScrollPane(preview1));
        preview2 = new JTable();
        previewPanel.add(new JScrollPane(preview2));
        centerSplit.setTopComponent(previewPanel);

        JSplitPane bottomSplit = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT);
        bottomSplit.setResizeWeight(0.4);
        centerSplit.setBottomComponent(bottomSplit);

        JPanel optionsPanel = createOptionsPanel();
        bottomSplit.setLeftComponent(optionsPanel);

        JPanel resultsPanel = createResultsPanel();
        bottomSplit.setRightComponent(resultsPanel);

        add(mainPanel);
    }

    private JPanel createFilePanel(int fileNum) {
        JPanel filePanel = new JPanel(new BorderLayout(5, 5));
        filePanel.setBorder(BorderFactory.createTitledBorder("File " + fileNum));

        JButton fileButton = new JButton("Choose File...");
        JComboBox<String> sheetComboBox = new JComboBox<>();
        sheetComboBox.setEnabled(false);
        JSpinner headerRowsSpinner = new JSpinner(new SpinnerNumberModel(1, 1, 10, 1));

        JPanel controlsPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 2, 0));
        controlsPanel.add(fileButton);
        controlsPanel.add(new JLabel("Headers:"));
        controlsPanel.add(headerRowsSpinner);
        controlsPanel.add(new JLabel("Sheet:"));

        JPanel topControls = new JPanel(new BorderLayout());
        topControls.add(controlsPanel, BorderLayout.WEST);
        topControls.add(sheetComboBox, BorderLayout.CENTER);

        JCheckBox chkNormalize = new JCheckBox("Enable Normalization (Transpose ticked rows)");
        chkNormalize.addActionListener(e -> {
            if (fileNum == 1) sheet1Changed(null); else sheet2Changed(null);
        });

        filePanel.add(topControls, BorderLayout.NORTH);
        filePanel.add(chkNormalize, BorderLayout.SOUTH);

        if (fileNum == 1) {
            sheet1ComboBox = sheetComboBox;
            headerRowsSpinner1 = headerRowsSpinner;
            chkNormalize1 = chkNormalize;
            fileButton.addActionListener(this::chooseFile1);
            sheet1ComboBox.addActionListener(this::sheet1Changed);
            headerRowsSpinner1.addChangeListener(e -> sheet1Changed(null));
        } else {
            sheet2ComboBox = sheetComboBox;
            headerRowsSpinner2 = headerRowsSpinner;
            chkNormalize2 = chkNormalize;
            fileButton.addActionListener(this::chooseFile2);
            sheet2ComboBox.addActionListener(this::sheet2Changed);
            headerRowsSpinner2.addChangeListener(e -> sheet2Changed(null));
        }
        return filePanel;
    }

    private JPanel createOptionsPanel() {
        JPanel optionsPanel = new JPanel(new BorderLayout(5, 5));
        optionsPanel.setBorder(BorderFactory.createTitledBorder("Comparison Options"));

        columnMappingPanel = new JPanel();
        columnMappingPanel.setLayout(new BoxLayout(columnMappingPanel, BoxLayout.Y_AXIS));
        JScrollPane mappingScrollPane = new JScrollPane(columnMappingPanel);

        JButton addMappingButton = new JButton("+ Add Key");
        addMappingButton.addActionListener(e -> addMappingRow());

        JPanel mappingContainer = new JPanel(new BorderLayout());
        mappingContainer.add(new JLabel("Map Key Columns:", SwingConstants.CENTER), BorderLayout.NORTH);
        mappingContainer.add(mappingScrollPane, BorderLayout.CENTER);
        mappingContainer.add(addMappingButton, BorderLayout.SOUTH);

        optionsPanel.add(mappingContainer, BorderLayout.CENTER);

        JPanel checkboxPanel = new JPanel(new GridLayout(3, 1));
        columnLevelComparison = new JCheckBox("Column-level comparison", false);
        detectMissingExtraRows = new JCheckBox("Detect missing/extra rows", true);
        checkboxPanel.add(columnLevelComparison);
        checkboxPanel.add(detectMissingExtraRows);
        optionsPanel.add(checkboxPanel, BorderLayout.WEST);

        addMappingRow();
        return optionsPanel;
    }

    private JPanel createResultsPanel() {
        JPanel resultsPanel = new JPanel(new BorderLayout());
        resultsPanel.setBorder(BorderFactory.createTitledBorder("Results"));
        resultsTable = new JTable();
        resultsPanel.add(new JScrollPane(resultsTable), BorderLayout.CENTER);

        JPanel actionPanel = new JPanel();
        JButton runButton = new JButton("Run Comparison");
        runButton.addActionListener(this::runComparison);
        JButton exportButton = new JButton("Export to Excel");
        exportButton.addActionListener(this::exportResults);
        JButton clearButton = new JButton("Clear Result");
        clearButton.addActionListener(this::clearResults);
        JButton exitButton = new JButton("Exit");
        exitButton.addActionListener(e -> System.exit(0));

        actionPanel.add(runButton);
        actionPanel.add(exportButton);
        actionPanel.add(clearButton);
        actionPanel.add(exitButton);
        resultsPanel.add(actionPanel, BorderLayout.SOUTH);
        return resultsPanel;
    }

    private void chooseFile1(ActionEvent e) {
        JFileChooser fileChooser = new JFileChooser();
        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            file1 = fileChooser.getSelectedFile();
            updatePostFileSelection(file1, sheet1ComboBox);
        }
    }

    private void chooseFile2(ActionEvent e) {
        JFileChooser fileChooser = new JFileChooser();
        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            file2 = fileChooser.getSelectedFile();
            updatePostFileSelection(file2, sheet2ComboBox);
        }
    }

    private void updatePostFileSelection(File file, JComboBox<String> sheetComboBox) {
        try {
            List<String> sheetNames = ExcelUtil.getSheetNames(file);
            sheetComboBox.setModel(new DefaultComboBoxModel<>(sheetNames.toArray(new String[0])));
            sheetComboBox.setEnabled(true);
            if (!sheetNames.isEmpty()) {
                // This will trigger sheetChanged event which calls loadSheetData
                sheetComboBox.setSelectedIndex(0);
            }
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this, "Error reading file: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void sheet1Changed(ActionEvent e) {
        String selectedSheet = (String) sheet1ComboBox.getSelectedItem();
        if (selectedSheet != null) {
            loadSheetData(1, selectedSheet);
        }
    }

    private void sheet2Changed(ActionEvent e) {
        String selectedSheet = (String) sheet2ComboBox.getSelectedItem();
        if (selectedSheet != null) {
            loadSheetData(2, selectedSheet);
        }
    }

    private void loadSheetData(int fileNum, String sheetName) {
        File file = (fileNum == 1) ? file1 : file2;
        JSpinner spinner = (fileNum == 1) ? headerRowsSpinner1 : headerRowsSpinner2;
        JCheckBox chkNormalize = (fileNum == 1) ? chkNormalize1 : chkNormalize2;
        if (file == null) return;

        int headerRows = (Integer) spinner.getValue();

        try {
            ExcelUtil.ExcelData data;
            if (chkNormalize.isSelected()) {
                data = ExcelUtil.readAndAutoNormalize(file, sheetName, headerRows);
            } else {
                data = ExcelUtil.readExcel(file, sheetName, headerRows, -1);
            }

            if (fileNum == 1) {
                data1 = data;
                updatePreviewTable(preview1, data1, 10);
            } else {
                data2 = data;
                updatePreviewTable(preview2, data2, 10);
            }
            updateColumnMappings();
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Error processing sheet '" + sheetName + "': " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void runComparison(ActionEvent e) {
        if (data1 == null || data2 == null) {
            JOptionPane.showMessageDialog(this, "Please load both Excel files and select sheets.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        Map<String, String> keyColumns = new HashMap<>();
        for (int i = 0; i < file1ColumnDropdowns.size(); i++) {
            String col1 = (String) file1ColumnDropdowns.get(i).getSelectedItem();
            String col2 = (String) file2ColumnDropdowns.get(i).getSelectedItem();
            if (col1 != null && col2 != null && !col1.isEmpty() && !col2.isEmpty()) {
                keyColumns.put(col1, col2);
            }
        }

        if (keyColumns.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please map at least one key column.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        lastResult = ComparatorLogic.compare(data1, data2, keyColumns, !columnLevelComparison.isSelected(), detectMissingExtraRows.isSelected());
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

    private void clearResults(ActionEvent e) {
        resultsTable.setModel(new DefaultTableModel());
        lastResult = null;
    }

    private void displayResults(ComparisonResult result) {
        List<String> displayHeaders = new ArrayList<>();
        displayHeaders.add("Status");

        List<String> headersForData;

        if (columnLevelComparison.isSelected()) {
            // New behavior: show only key columns
            headersForData = new ArrayList<>();
            for (int i = 0; i < file1ColumnDropdowns.size(); i++) {
                String col1 = (String) file1ColumnDropdowns.get(i).getSelectedItem();
                String col2 = (String) file2ColumnDropdowns.get(i).getSelectedItem();
                if (col1 != null && !col1.isEmpty()) {
                    headersForData.add(col1);
                }
                 if (col2 != null && !col2.isEmpty()) {
                    headersForData.add(col2);
                }
            }
            // Remove duplicates
            headersForData = new ArrayList<>(new LinkedHashSet<>(headersForData));
        } else {
            // Default behavior: show all columns
            headersForData = result.getHeaders();
        }
        displayHeaders.addAll(headersForData);

        DefaultTableModel model = new DefaultTableModel(displayHeaders.toArray(new String[0]), 0);

        List<String> headers1 = data1.getHeaders();
        List<String> headers2 = data2.getHeaders();

        for (ComparisonResult.ResultRow row : result.getResultRows()) {
            List<Object> rowData = new ArrayList<>();
            rowData.add(row.getStatus());
            for (String header : headersForData) {
                Object val;

                List<Integer> indices1 = getAllIndices(headers1, header);
                List<Integer> indices2 = getAllIndices(headers2, header);

                switch (row.getStatus()) {
                    case MISSING_IN_FILE_1:
                        val = getCombinedValue(row.getData2(), indices2);
                        break;
                    case MISSING_IN_FILE_2:
                        val = getCombinedValue(row.getData1(), indices1);
                        break;
                    case MATCH:
                         val = getCombinedValue(row.getData1(), indices1);
                        break;
                    case MISMATCH:
                        Object val1 = getCombinedValue(row.getData1(), indices1);
                        Object val2 = getCombinedValue(row.getData2(), indices2);
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
        resultsTable.setDefaultRenderer(Object.class, new ResultCellRenderer(result, data1, data2));
    }

    private List<Integer> getAllIndices(List<String> headers, String header) {
        List<Integer> indices = new ArrayList<>();
        if (headers == null) return indices;
        for (int i = 0; i < headers.size(); i++) {
            if (header.equals(headers.get(i))) {
                indices.add(i);
            }
        }
        return indices;
    }

    private Object getCombinedValue(List<Object> data, List<Integer> indices) {
        if (indices.isEmpty() || data == null) return "";
        if (indices.size() == 1) {
            int idx = indices.get(0);
            return (idx < data.size()) ? data.get(idx) : "";
        }
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < indices.size(); i++) {
            int idx = indices.get(i);
            if (idx < data.size()) {
                sb.append(data.get(idx));
            }
            if (i < indices.size() - 1) {
                sb.append(" | ");
            }
        }
        return sb.toString();
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
        String[] headers1 = (data1 != null) ? new LinkedHashSet<>(data1.getHeaders()).toArray(new String[0]) : new String[0];
        String[] headers2 = (data2 != null) ? new LinkedHashSet<>(data2.getHeaders()).toArray(new String[0]) : new String[0];

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
    private final ExcelUtil.ExcelData data1, data2;

    public ResultCellRenderer(ComparisonResult result, ExcelUtil.ExcelData data1, ExcelUtil.ExcelData data2) {
        this.result = result;
        this.data1 = data1;
        this.data2 = data2;
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
                        String header = table.getColumnName(column);
                        Object val1 = getCombinedValue(resultRow.getData1(), getAllIndices(data1.getHeaders(), header));
                        Object val2 = getCombinedValue(resultRow.getData2(), getAllIndices(data2.getHeaders(), header));
                         if (!java.util.Objects.equals(val1, val2)) {
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

    private List<Integer> getAllIndices(List<String> headers, String header) {
        List<Integer> indices = new ArrayList<>();
        if (headers == null) return indices;
        for (int i = 0; i < headers.size(); i++) {
            if (header.equals(headers.get(i))) {
                indices.add(i);
            }
        }
        return indices;
    }

    private Object getCombinedValue(List<Object> data, List<Integer> indices) {
        if (indices.isEmpty() || data == null) return "";
        if (indices.size() == 1) {
            int idx = indices.get(0);
            return (idx < data.size()) ? data.get(idx) : "";
        }
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < indices.size(); i++) {
            int idx = indices.get(i);
            if (idx < data.size()) {
                sb.append(data.get(idx));
            }
            if (i < indices.size() - 1) {
                sb.append(" | ");
            }
        }
        return sb.toString();
    }
}
