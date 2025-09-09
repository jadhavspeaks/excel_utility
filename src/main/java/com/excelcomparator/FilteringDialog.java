package com.excelcomparator;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class FilteringDialog extends JDialog {

    private CardLayout cardLayout;
    private JPanel mainPanel;
    private File file1, file2;
    private JTextField file1Path, file2Path;
    private JRadioButton f1toF2Radio;

    private JPanel attributeMappingPanel, reportMappingPanel, productMappingPanel;
    private JComboBox<String> file1AttributeCombo, file2ReportCombo, file2ProductCombo;

    private List<String> headers1, headers2;
    private JTable resultsTable;
    private ExcelUtil.ExcelData filteredResult;

    private static final String LOAD_PANEL = "LoadPanel";
    private static final String MAPPING_PANEL = "MappingPanel";
    private static final String RESULTS_PANEL = "ResultsPanel";

    public FilteringDialog(Frame owner) {
        super(owner, "Advanced Filtering Tool", false);
        setSize(800, 600);
        setLocationRelativeTo(owner);

        cardLayout = new CardLayout();
        mainPanel = new JPanel(cardLayout);

        mainPanel.add(createLoadPanel(), LOAD_PANEL);
        mainPanel.add(createMappingPanel(), MAPPING_PANEL);
        mainPanel.add(createResultsPanel(), RESULTS_PANEL);

        add(mainPanel);
    }

    private JPanel createLoadPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        gbc.gridx = 0; gbc.gridy = 0;
        panel.add(new JLabel("Filter Definition File (Excel 1):"), gbc);
        file1Path = new JTextField(40);
        file1Path.setEditable(false);
        gbc.gridx = 1; panel.add(file1Path, gbc);
        JButton chooseFile1Button = new JButton("Choose...");
        chooseFile1Button.addActionListener(e -> chooseFile(1, file1Path));
        gbc.gridx = 2; panel.add(chooseFile1Button, gbc);

        gbc.gridx = 0; gbc.gridy = 1;
        panel.add(new JLabel("Data File (Excel 2):"), gbc);
        file2Path = new JTextField(40);
        file2Path.setEditable(false);
        gbc.gridx = 1; panel.add(file2Path, gbc);
        JButton chooseFile2Button = new JButton("Choose...");
        chooseFile2Button.addActionListener(e -> chooseFile(2, file2Path));
        gbc.gridx = 2; panel.add(chooseFile2Button, gbc);

        gbc.gridy = 3; gbc.gridwidth = 1;
        panel.add(new JLabel("Filter Direction:"), gbc);
        f1toF2Radio = new JRadioButton("Use Excel 1 to filter Excel 2", true);
        JRadioButton f2toF1Radio = new JRadioButton("Use Excel 2 to filter Excel 1");
        ButtonGroup directionGroup = new ButtonGroup();
        directionGroup.add(f1toF2Radio);
        directionGroup.add(f2toF1Radio);
        JPanel radioPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        radioPanel.add(f1toF2Radio);
        radioPanel.add(f2toF1Radio);
        gbc.gridx = 1; gbc.gridwidth = 2; panel.add(radioPanel, gbc);

        gbc.gridy = 4; gbc.gridx = 1; gbc.gridwidth = 1;
        JButton nextButton = new JButton("Next ->");
        nextButton.addActionListener(e -> showMappingPanel());
        panel.add(nextButton, gbc);

        return panel;
    }

    private JPanel createMappingPanel() {
        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBorder(BorderFactory.createTitledBorder("Step 2: Map Columns"));
        JPanel mappingSectionsPanel = new JPanel();
        mappingSectionsPanel.setLayout(new BoxLayout(mappingSectionsPanel, BoxLayout.Y_AXIS));

        attributeMappingPanel = createMappingSection("Attribute Mapping (1 to Many)");
        reportMappingPanel = createMappingSection("Report Mapping (Many to 1)");
        productMappingPanel = createMappingSection("Product Mapping (Many to 1)");

        mappingSectionsPanel.add(attributeMappingPanel);
        mappingSectionsPanel.add(reportMappingPanel);
        mappingSectionsPanel.add(productMappingPanel);
        panel.add(new JScrollPane(mappingSectionsPanel), BorderLayout.CENTER);

        JPanel navPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        JButton backButton = new JButton("<- Back");
        backButton.addActionListener(e -> cardLayout.show(mainPanel, LOAD_PANEL));
        JButton runButton = new JButton("Run Filter");
        runButton.addActionListener(e -> runFilter());
        navPanel.add(backButton);
        navPanel.add(runButton);
        panel.add(navPanel, BorderLayout.SOUTH);

        return panel;
    }

    private JPanel createResultsPanel() {
        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBorder(BorderFactory.createTitledBorder("Step 3: Filtered Results"));
        resultsTable = new JTable();
        panel.add(new JScrollPane(resultsTable), BorderLayout.CENTER);

        JPanel navPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        JButton backButton = new JButton("<- Back");
        backButton.addActionListener(e -> cardLayout.show(mainPanel, MAPPING_PANEL));
        JButton exportButton = new JButton("Export to Excel");
        exportButton.addActionListener(e -> exportFilteredData());
        navPanel.add(backButton);
        navPanel.add(exportButton);
        panel.add(navPanel, BorderLayout.SOUTH);

        return panel;
    }

    private void chooseFile(int fileNum, JTextField pathField) {
        JFileChooser fileChooser = new JFileChooser();
        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            if (fileNum == 1) file1 = selectedFile;
            else file2 = selectedFile;
            pathField.setText(selectedFile.getAbsolutePath());
        }
    }

    private void showMappingPanel() {
        if (file1 == null || file2 == null) {
            JOptionPane.showMessageDialog(this, "Please select both files.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }
        try {
            headers1 = ExcelUtil.readExcel(file1, null, 1, 1).getHeaders();
            headers2 = ExcelUtil.readExcel(file2, null, 1, 1).getHeaders();
            populateMappingPanel();
            cardLayout.show(mainPanel, MAPPING_PANEL);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Error reading file headers: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private JPanel createMappingSection(String title) {
        JPanel section = new JPanel();
        section.setLayout(new BoxLayout(section, BoxLayout.Y_AXIS));
        section.setBorder(BorderFactory.createTitledBorder(title));
        return section;
    }

    private void populateMappingPanel() {
        attributeMappingPanel.removeAll();
        JPanel attr1Panel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        attr1Panel.add(new JLabel("Attribute in File 1:"));
        file1AttributeCombo = new JComboBox<>(headers1.toArray(new String[0]));
        attr1Panel.add(file1AttributeCombo);
        attributeMappingPanel.add(attr1Panel);
        JButton addAttrButton = new JButton("Add Attribute from File 2");
        attributeMappingPanel.add(addAttrButton);
        addAttrButton.addActionListener(e -> addDynamicMappingRow(attributeMappingPanel, "Attribute in File 2:", headers2));
        addDynamicMappingRow(attributeMappingPanel, "Attribute in File 2:", headers2);

        reportMappingPanel.removeAll();
        JPanel report2Panel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        report2Panel.add(new JLabel("Report Name in File 2:"));
        file2ReportCombo = new JComboBox<>(headers2.toArray(new String[0]));
        report2Panel.add(file2ReportCombo);
        reportMappingPanel.add(report2Panel);
        JButton addReportButton = new JButton("Add Report Column from File 1");
        reportMappingPanel.add(addReportButton);
        addReportButton.addActionListener(e -> addDynamicMappingRow(reportMappingPanel, "Report Column in File 1:", headers1));
        addDynamicMappingRow(reportMappingPanel, "Report Column in File 1:", headers1);

        productMappingPanel.removeAll();
        JPanel product2Panel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        product2Panel.add(new JLabel("Product in File 2:"));
        file2ProductCombo = new JComboBox<>(headers2.toArray(new String[0]));
        product2Panel.add(file2ProductCombo);
        productMappingPanel.add(product2Panel);
        JButton addProductButton = new JButton("Add Product Column from File 1");
        productMappingPanel.add(addProductButton);
        addProductButton.addActionListener(e -> addDynamicMappingRow(productMappingPanel, "Product Column in File 1:", headers1));
        addDynamicMappingRow(productMappingPanel, "Product Column in File 1:", headers1);

        mainPanel.revalidate();
        mainPanel.repaint();
    }

    private void addDynamicMappingRow(JPanel parent, String label, List<String> headers) {
        JPanel row = new JPanel(new FlowLayout(FlowLayout.LEFT));
        row.add(new JLabel(label));
        row.add(new JComboBox<>(headers.toArray(new String[0])));
        JButton removeButton = new JButton("X");
        removeButton.addActionListener(e -> {
            parent.remove(row);
            parent.revalidate();
            parent.repaint();
        });
        row.add(removeButton);
        parent.add(row, parent.getComponentCount() - 1);
        parent.revalidate();
    }

    private void runFilter() {
        FilterMappingConfig config = gatherMappingConfig();
        if (config == null) return;

        try {
            ExcelUtil.ExcelData data1 = ExcelUtil.readExcel(file1, null, 1, -1);
            ExcelUtil.ExcelData data2 = ExcelUtil.readExcel(file2, null, 1, -1);
            filteredResult = FilterLogic.filter(data1, data2, config);

            if (filteredResult.getData().isEmpty()) {
                resultsTable.setModel(new DefaultTableModel(new Object[][]{{"No data found for given filters"}}, new Object[]{"Status"}));
            } else {
                DefaultTableModel model = new DefaultTableModel(
                    filteredResult.getData().stream().map(List::toArray).toArray(Object[][]::new),
                    filteredResult.getHeaders().toArray()
                );
                resultsTable.setModel(model);
            }
            cardLayout.show(mainPanel, RESULTS_PANEL);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Error running filter: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private FilterMappingConfig gatherMappingConfig() {
        FilterMappingConfig config = new FilterMappingConfig();
        config.setDirection(f1toF2Radio.isSelected() ? FilterMappingConfig.FilterDirection.FILE1_FILTERS_FILE2 : FilterMappingConfig.FilterDirection.FILE2_FILTERS_FILE1);

        config.setFile1_AttributeColumn((String) file1AttributeCombo.getSelectedItem());
        config.setFile2_AttributeColumns(getDynamicMappings(attributeMappingPanel));

        config.setFile2_ReportColumn((String) file2ReportCombo.getSelectedItem());
        config.setFile1_ReportColumns(getDynamicMappings(reportMappingPanel));

        config.setFile2_ProductColumn((String) file2ProductCombo.getSelectedItem());
        config.setFile1_ProductColumns(getDynamicMappings(productMappingPanel));

        return config;
    }

    private List<String> getDynamicMappings(JPanel parent) {
        List<String> mappings = new ArrayList<>();
        for (Component comp : parent.getComponents()) {
            if (comp instanceof JPanel) {
                for (Component subComp : ((JPanel) comp).getComponents()) {
                    if (subComp instanceof JComboBox) {
                        mappings.add((String) ((JComboBox<?>) subComp).getSelectedItem());
                    }
                }
            }
        }
        return mappings;
    }

    private void exportFilteredData() {
        if (filteredResult == null || filteredResult.getData().isEmpty()) {
            JOptionPane.showMessageDialog(this, "No data to export.", "Info", JOptionPane.INFORMATION_MESSAGE);
            return;
        }
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Save Filtered Data");
        if (fileChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            File fileToSave = fileChooser.getSelectedFile();
            if (!fileToSave.getName().toLowerCase().endsWith(".xlsx")) {
                fileToSave = new File(fileToSave.getParentFile(), fileToSave.getName() + ".xlsx");
            }
            try {
                ExcelUtil.writeFilteredDataToExcel(filteredResult, fileToSave);
                JOptionPane.showMessageDialog(this, "Export successful!", "Success", JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(this, "Error exporting file: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }
}
