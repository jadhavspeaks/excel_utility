package com.excelcomparator;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class SimpleFilterDialog extends JDialog {

    private File inputFile;
    private JTextField inputFilePath;
    private JComboBox<String> sheetCombo;
    private JSpinner headerRowsSpinner;
    private JTextArea logArea;
    private JButton startButton;
    private JPanel conditionsPanel;
    private List<FilterConditionRow> conditionRows = new ArrayList<>();


    public SimpleFilterDialog(Frame owner) {
        super(owner, "Multi-Condition Filter", true);
        setSize(800, 700);
        setLocationRelativeTo(owner);
        setLayout(new BorderLayout(10, 10));

        add(createConfigPanel(), BorderLayout.NORTH);
        add(createLogPanel(), BorderLayout.CENTER);
        add(createButtonPanel(), BorderLayout.SOUTH);
    }

    private JPanel createConfigPanel() {
        JPanel panel = new JPanel(new BorderLayout(10,10));
        panel.setBorder(BorderFactory.createTitledBorder("Configuration"));

        JPanel topPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        // Input File
        gbc.gridx = 0; gbc.gridy = 0;
        topPanel.add(new JLabel("Input Excel File:"), gbc);
        inputFilePath = new JTextField(30);
        inputFilePath.setEditable(false);
        gbc.gridx = 1; gbc.gridwidth = 1; gbc.weightx=1; topPanel.add(inputFilePath, gbc);
        JButton chooseFileButton = new JButton("Choose...");
        chooseFileButton.addActionListener(e -> chooseInputFile());
        gbc.gridx = 2; gbc.gridwidth = 1; gbc.weightx=0; topPanel.add(chooseFileButton, gbc);

        // Header Row Selection
        gbc.gridx = 0; gbc.gridy = 1;
        topPanel.add(new JLabel("Header Rows:"), gbc);
        headerRowsSpinner = new JSpinner(new SpinnerNumberModel(1, 1, 100, 1));
        headerRowsSpinner.addChangeListener(e -> loadColumnsForSheet());
        gbc.gridx = 1; topPanel.add(headerRowsSpinner, gbc);

        // Sheet Selection
        gbc.gridx = 0; gbc.gridy = 2;
        topPanel.add(new JLabel("Sheet to Filter:"), gbc);
        sheetCombo = new JComboBox<>();
        sheetCombo.addActionListener(e -> loadColumnsForSheet());
        gbc.gridx = 1; gbc.gridwidth = 2; topPanel.add(sheetCombo, gbc);

        panel.add(topPanel, BorderLayout.NORTH);

        // Conditions Panel
        conditionsPanel = new JPanel();
        conditionsPanel.setLayout(new BoxLayout(conditionsPanel, BoxLayout.Y_AXIS));
        conditionsPanel.setBorder(BorderFactory.createTitledBorder("Filter Conditions (AND)"));
        panel.add(new JScrollPane(conditionsPanel), BorderLayout.CENTER);

        JButton addConditionButton = new JButton("Add Condition");
        addConditionButton.addActionListener(e -> addConditionRow());
        panel.add(addConditionButton, BorderLayout.SOUTH);

        addConditionRow(); // Start with one condition

        return panel;
    }

    private JScrollPane createLogPanel() {
        logArea = new JTextArea();
        logArea.setEditable(false);
        logArea.setBorder(BorderFactory.createTitledBorder("Log"));
        return new JScrollPane(logArea);
    }

    private JPanel createButtonPanel() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        startButton = new JButton("Start Filtering");
        startButton.addActionListener(e -> startFiltering());
        panel.add(startButton);
        JButton closeButton = new JButton("Close");
        closeButton.addActionListener(e -> setVisible(false));
        panel.add(closeButton);
        return panel;
    }

    private void chooseInputFile() {
        JFileChooser fileChooser = new JFileChooser();
        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            inputFile = fileChooser.getSelectedFile();
            inputFilePath.setText(inputFile.getAbsolutePath());
            loadSheets();
        }
    }

    private void loadSheets() {
        if (inputFile == null) return;
        SwingWorker<List<String>, Void> worker = new SwingWorker<>() {
            @Override
            protected List<String> doInBackground() throws Exception {
                return ExcelUtil.getSheetNames(inputFile);
            }

            @Override
            protected void done() {
                try {
                    List<String> sheetNames = get();
                    sheetCombo.setModel(new DefaultComboBoxModel<>(sheetNames.toArray(new String[0])));
                    if (!sheetNames.isEmpty()) {
                        loadColumnsForSheet();
                    }
                } catch (Exception e) {
                    log("Error loading sheets: " + e.getMessage());
                }
            }
        };
        worker.execute();
    }

    private void addConditionRow() {
        FilterConditionRow newRow = new FilterConditionRow(new ArrayList<>());
        JButton removeButton = new JButton("Remove");
        newRow.add(removeButton);
        removeButton.addActionListener(e -> {
            conditionsPanel.remove(newRow);
            conditionRows.remove(newRow);
            conditionsPanel.revalidate();
            conditionsPanel.repaint();
        });
        conditionRows.add(newRow);
        conditionsPanel.add(newRow);
        conditionsPanel.revalidate();
        conditionsPanel.repaint();
        loadColumnsForSheet(); // To populate the new row with columns if already loaded
    }

    private void loadColumnsForSheet() {
        String selectedSheet = (String) sheetCombo.getSelectedItem();
        if (inputFile == null || selectedSheet == null) return;

        int headerRows = (int) headerRowsSpinner.getValue();

        SwingWorker<List<String>, Void> worker = new SwingWorker<>() {
            @Override
            protected List<String> doInBackground() throws Exception {
                return ExcelUtil.getHeaders(inputFile, selectedSheet, headerRows);
            }

            @Override
            protected void done() {
                try {
                    List<String> headers = get();
                    for (FilterConditionRow row : conditionRows) {
                        row.setColumnNames(headers);
                    }
                } catch (Exception e) {
                    log("Error loading columns: " + e.getMessage());
                }
            }
        };
        worker.execute();
    }

    private void startFiltering() {
        if (inputFile == null || sheetCombo.getSelectedItem() == null) {
            JOptionPane.showMessageDialog(this, "Please select an input file and sheet.", "Configuration Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        List<FilterCondition> conditions = new ArrayList<>();
        for (FilterConditionRow row : conditionRows) {
            if (row.getValue().isBlank()) {
                JOptionPane.showMessageDialog(this, "Please ensure all filter condition values are filled in.", "Configuration Error", JOptionPane.ERROR_MESSAGE);
                return;
            }
            conditions.add(new FilterCondition(row.getSelectedColumn(), row.getSelectedOperator(), row.getValue()));
        }

        if (conditions.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please add at least one filter condition.", "Configuration Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        String outputName = "filtered_" + inputFile.getName();
        File outputFile = new File(inputFile.getParent(), outputName);

        String sheetName = (String) sheetCombo.getSelectedItem();
        int headerRows = (int) headerRowsSpinner.getValue();

        log("Starting filter process...");
        log("Output will be saved to: " + outputFile.getAbsolutePath());
        startButton.setEnabled(false);

        new SwingWorker<Void, String>() {
            @Override
            protected Void doInBackground() throws Exception {
                ExcelFilter filter = new ExcelFilter();
                filter.processFile(inputFile, outputFile, sheetName, headerRows, conditions, this::publish);
                return null;
            }

            @Override
            protected void process(List<String> chunks) {
                for (String message : chunks) {
                    log(message);
                }
            }

            @Override
            protected void done() {
                try {
                    get(); // To catch exceptions from doInBackground
                    log("Filter process completed successfully.");
                    String message = "Filtering complete!\nFile saved at: " + outputFile.getAbsolutePath();
                    JTextArea textArea = new JTextArea(message);
                    textArea.setEditable(false);
                    JScrollPane scrollPane = new JScrollPane(textArea);
                    scrollPane.setPreferredSize(new Dimension(400, 100));
                    JOptionPane.showMessageDialog(SimpleFilterDialog.this, scrollPane, "Success", JOptionPane.INFORMATION_MESSAGE);
                } catch (Exception e) {
                    log("Error during filtering: " + e.getMessage());
                    e.printStackTrace();
                    JOptionPane.showMessageDialog(SimpleFilterDialog.this, "An error occurred: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                } finally {
                    startButton.setEnabled(true);
                }
            }
        }.execute();
    }

    private void log(String message) {
        logArea.append(message + "\n");
    }
}
