package com.excelcomparator;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;

public class ExcelComparator extends JFrame {

    private JTable preview1, preview2, resultsTable;
    private ExcelUtil.ExcelData data1, data2;
    private File file1, file2;
    private JComboBox<String> sheet1ComboBox, sheet2ComboBox;
    private JSpinner headerRowsSpinner1, headerRowsSpinner2;
    private JCheckBox chkNormalize1, chkNormalize2;
    private JTextArea summaryArea;
    private ComparisonResult lastResult;
    private JPanel columnMappingPanel;
    private List<JComboBox<String>> file1ColumnDropdowns = new ArrayList<>();
    private List<JComboBox<String>> file2ColumnDropdowns = new ArrayList<>();
    private JCheckBox columnLevelComparison;
    private JCheckBox detectMissingExtraRows;

    public ExcelComparator() {
        setTitle("Excel Comparator");
        setSize(1280, 800);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);

        JPanel mainPanel = new JPanel(new BorderLayout(5, 5));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));

        JPanel topPanel = new JPanel(new GridLayout(1, 2, 5, 5));
        topPanel.add(createFilePanel(1));
        topPanel.add(createFilePanel(2));
        mainPanel.add(topPanel, BorderLayout.NORTH);

        JSplitPane centerSplit = new JSplitPane(JSplitPane.VERTICAL_SPLIT);
        centerSplit.setResizeWeight(0.5);
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

        bottomSplit.setLeftComponent(createOptionsPanel());
        bottomSplit.setRightComponent(createResultsPanel());

        add(mainPanel);

        JMenuBar menuBar = new JMenuBar();
        JMenu toolsMenu = new JMenu("Tools");

        JMenuItem simpleNormalizeItem = new JMenuItem("Simple Normalizer");
        simpleNormalizeItem.addActionListener(e -> {
            SimpleNormalizerDialog dialog = new SimpleNormalizerDialog(this);
            dialog.setVisible(true);
        });
        toolsMenu.add(simpleNormalizeItem);

        JMenuItem simpleFilterItem = new JMenuItem("Simple Keyword Filter");
        simpleFilterItem.addActionListener(e -> {
            SimpleFilterDialog dialog = new SimpleFilterDialog(this);
            dialog.setVisible(true);
        });
        toolsMenu.add(simpleFilterItem);

        menuBar.add(toolsMenu);
        setJMenuBar(menuBar);
    }

    private JPanel createFilePanel(int fileNum) {
        JPanel panel = new JPanel(new BorderLayout(5, 5));
        panel.setBorder(BorderFactory.createTitledBorder("File " + fileNum));

        JButton fileButton = new JButton("Choose File...");
        JComboBox<String> sheetComboBox = new JComboBox<>();
        sheetComboBox.setEnabled(false);
        JSpinner headerRowsSpinner = new JSpinner(new SpinnerNumberModel(1, 1, 10, 1));

        fileButton.addActionListener(e -> chooseFile(fileNum, sheetComboBox, headerRowsSpinner));
        sheetComboBox.addActionListener(e -> loadSheetData(fileNum, (String) sheetComboBox.getSelectedItem(), (int) headerRowsSpinner.getValue()));
        headerRowsSpinner.addChangeListener(e -> loadSheetData(fileNum, (String) sheetComboBox.getSelectedItem(), (int) headerRowsSpinner.getValue()));

        JPanel controlsPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        controlsPanel.add(fileButton);
        controlsPanel.add(new JLabel("Headers:"));
        controlsPanel.add(headerRowsSpinner);
        controlsPanel.add(new JLabel("Sheet:"));
        controlsPanel.add(sheetComboBox);
        panel.add(controlsPanel, BorderLayout.NORTH);

        JCheckBox chkNormalize = new JCheckBox("Enable Automated Normalization");
        chkNormalize.addActionListener(e -> loadSheetData(fileNum, (String) sheetComboBox.getSelectedItem(), (int) headerRowsSpinner.getValue()));
        panel.add(chkNormalize, BorderLayout.SOUTH);

        if (fileNum == 1) {
            sheet1ComboBox = sheetComboBox;
            headerRowsSpinner1 = headerRowsSpinner;
            chkNormalize1 = chkNormalize;
        } else {
            sheet2ComboBox = sheetComboBox;
            headerRowsSpinner2 = headerRowsSpinner;
            chkNormalize2 = chkNormalize;
        }
        return panel;
    }

    private JPanel createOptionsPanel() {
        JPanel panel = new JPanel(new BorderLayout(5, 5));
        panel.setBorder(BorderFactory.createTitledBorder("Comparison Options"));

        columnMappingPanel = new JPanel();
        columnMappingPanel.setLayout(new BoxLayout(columnMappingPanel, BoxLayout.Y_AXIS));
        JScrollPane mappingScrollPane = new JScrollPane(columnMappingPanel);
        addMappingRow();

        JButton addMappingButton = new JButton("+ Add Key");
        addMappingButton.addActionListener(e -> addMappingRow());

        JPanel mappingContainer = new JPanel(new BorderLayout());
        mappingContainer.add(new JLabel("Map Key Columns:", SwingConstants.CENTER), BorderLayout.NORTH);
        mappingContainer.add(mappingScrollPane, BorderLayout.CENTER);
        mappingContainer.add(addMappingButton, BorderLayout.SOUTH);
        panel.add(mappingContainer, BorderLayout.CENTER);

        JPanel checkboxPanel = new JPanel(new GridLayout(3, 1));
        columnLevelComparison = new JCheckBox("Show All Columns", false);
        detectMissingExtraRows = new JCheckBox("Detect missing/extra rows", true);
        checkboxPanel.add(columnLevelComparison);
        checkboxPanel.add(detectMissingExtraRows);
        panel.add(checkboxPanel, BorderLayout.WEST);

        return panel;
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
        row.add(new JLabel("File 1:"));
        row.add(cb1);
        row.add(new JLabel(" vs File 2:"));
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
            String selected = (String) cb.getSelectedItem();
            cb.setModel(new DefaultComboBoxModel<>(headers1));
            cb.setSelectedItem(selected);
        }
        for (JComboBox<String> cb : file2ColumnDropdowns) {
            String selected = (String) cb.getSelectedItem();
            cb.setModel(new DefaultComboBoxModel<>(headers2));
            cb.setSelectedItem(selected);
        }
    }

    private JPanel createResultsPanel() {
        JPanel panel = new JPanel(new BorderLayout(5, 5));
        panel.setBorder(BorderFactory.createTitledBorder("Results"));

        summaryArea = new JTextArea(8, 80);
        summaryArea.setEditable(false);
        summaryArea.setFont(new Font("Monospaced", Font.PLAIN, 12));
        panel.add(new JScrollPane(summaryArea), BorderLayout.NORTH);

        resultsTable = new JTable();
        panel.add(new JScrollPane(resultsTable), BorderLayout.CENTER);

        JPanel actionPanel = new JPanel();
        JButton runButton = new JButton("Run Comparison");
        runButton.addActionListener(e -> runComparison());
        JButton exportButton = new JButton("Export to Excel");
        exportButton.addActionListener(e -> exportResults());
        JButton clearButton = new JButton("Clear Result");
        clearButton.addActionListener(e -> clearResults());
        JButton exitButton = new JButton("Exit");
        exitButton.addActionListener(e -> System.exit(0));
        actionPanel.add(runButton);
        actionPanel.add(exportButton);
        actionPanel.add(clearButton);
        actionPanel.add(exitButton);
        panel.add(actionPanel, BorderLayout.SOUTH);
        return panel;
    }

    private void chooseFile(int fileNum, JComboBox<String> sheetComboBox, JSpinner headerSpinner) {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xls", "xlsx"));
        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            if (fileNum == 1) file1 = selectedFile;
            else file2 = selectedFile;
            try {
                List<String> sheetNames = ExcelUtil.getSheetNames(selectedFile);
                sheetComboBox.setModel(new DefaultComboBoxModel<>(sheetNames.toArray(new String[0])));
                sheetComboBox.setEnabled(true);
                if (!sheetNames.isEmpty()) {
                    loadSheetData(fileNum, sheetNames.get(0), (int) headerSpinner.getValue());
                }
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(this, "Error reading file: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void loadSheetData(int fileNum, String sheetName, int headerRows) {
        File file = (fileNum == 1) ? file1 : file2;
        JCheckBox chkNormalize = (fileNum == 1) ? chkNormalize1 : chkNormalize2;
        if (file == null || sheetName == null) return;
        try {
            ExcelUtil.ExcelData data = (chkNormalize != null && chkNormalize.isSelected())
                ? ExcelUtil.readAndAutoNormalize(file, sheetName, headerRows)
                : ExcelUtil.readExcel(file, sheetName, headerRows, 10);
            if (fileNum == 1) data1 = data;
            else data2 = data;
            updatePreviewTable((fileNum == 1) ? preview1 : preview2, data);
            updateColumnMappings();
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this, "Error loading sheet data: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void updatePreviewTable(JTable table, ExcelUtil.ExcelData excelData) {
        if (excelData == null) return;
        DefaultTableModel model = new DefaultTableModel(
            excelData.getData().stream().map(List::toArray).toArray(Object[][]::new),
            excelData.getHeaders().toArray()
        );
        table.setModel(model);
    }

    private void runComparison() {
        if (data1 == null || data2 == null) {
            JOptionPane.showMessageDialog(this, "Please load both files first.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }
        Map<String, String> keyColumns = new HashMap<>();
        for (int i = 0; i < file1ColumnDropdowns.size(); i++) {
            String col1 = (String) file1ColumnDropdowns.get(i).getSelectedItem();
            String col2 = (String) file2ColumnDropdowns.get(i).getSelectedItem();
            if (col1 != null && !col1.isEmpty() && col2 != null && !col2.isEmpty()) {
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

    private void displayResults(ComparisonResult result) {
        summaryArea.setText(result.getSummary().toString());
        summaryArea.setCaretPosition(0);

        List<String> displayHeaders = new ArrayList<>();
        displayHeaders.add("File 1 Row");
        displayHeaders.add("File 2 Row");
        displayHeaders.add("Status");
        displayHeaders.add("Mismatched Columns");

        List<String> headersForData;
        if (columnLevelComparison.isSelected()) {
            headersForData = result.getHeaders();
        } else {
            headersForData = new ArrayList<>();
            for (int i = 0; i < file1ColumnDropdowns.size(); i++) {
                String col1 = (String) file1ColumnDropdowns.get(i).getSelectedItem();
                String col2 = (String) file2ColumnDropdowns.get(i).getSelectedItem();
                if (col1 != null && !col1.isEmpty()) headersForData.add(col1);
                if (col2 != null && !col2.isEmpty()) headersForData.add(col2);
            }
            headersForData = new ArrayList<>(new LinkedHashSet<>(headersForData));
        }
        displayHeaders.addAll(headersForData);

        DefaultTableModel model = new DefaultTableModel(displayHeaders.toArray(new String[0]), 0);
        List<String> headers1 = data1.getHeaders();
        List<String> headers2 = data2.getHeaders();

        for (ComparisonResult.ResultRow row : result.getResultRows()) {
            List<Object> rowData = new ArrayList<>();
            rowData.add(row.getOriginalRowNum1() > 0 ? row.getOriginalRowNum1() : "");
            rowData.add(row.getOriginalRowNum2() > 0 ? row.getOriginalRowNum2() : "");
            rowData.add(row.getStatus());
            rowData.add(String.join(", ", row.getMismatchedColumns()));
            for (String header : headersForData) {
                Object val1 = ExcelUtil.getCombinedValue(row.getData1(), ExcelUtil.getAllIndices(headers1, header));
                Object val2 = ExcelUtil.getCombinedValue(row.getData2(), ExcelUtil.getAllIndices(headers2, header));
                if (row.getStatus() == ComparisonResult.RowStatus.MISMATCH && !Objects.equals(val1, val2)) {
                    rowData.add(String.format("%s -> %s", val1, val2));
                } else if (row.getStatus() == ComparisonResult.RowStatus.MISSING_IN_FILE_1) {
                    rowData.add(val2);
                } else {
                    rowData.add(val1);
                }
            }
            model.addRow(rowData.toArray());
        }
        resultsTable.setModel(model);
    }

    private void exportResults() {
        if (lastResult == null) {
            JOptionPane.showMessageDialog(this, "No results to export.", "Info", JOptionPane.INFORMATION_MESSAGE);
            return;
        }
        JFileChooser fileChooser = new JFileChooser();
        if (fileChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            File fileToSave = fileChooser.getSelectedFile();
            if (!fileToSave.getName().toLowerCase().endsWith(".xlsx")) {
                fileToSave = new File(fileToSave.getParentFile(), fileToSave.getName() + ".xlsx");
            }
            try {
                ExcelUtil.writeResultToExcel(lastResult, data1, data2, fileToSave);
                JOptionPane.showMessageDialog(this, "Export successful!", "Success", JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(this, "Error exporting file: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void clearResults() {
        summaryArea.setText("");
        resultsTable.setModel(new DefaultTableModel());
        lastResult = null;
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            ExcelComparator app = new ExcelComparator();
            app.setVisible(true);
        });
    }
}
