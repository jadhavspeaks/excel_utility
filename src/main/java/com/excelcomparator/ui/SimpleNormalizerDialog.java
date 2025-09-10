package com.excelcomparator.ui;

import com.excelcomparator.logic.StreamingExcelProcessor;
import com.excelcomparator.util.ExcelUtil;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.util.List;

public class SimpleNormalizerDialog extends JDialog {

    private File inputFile;
    private JTextField inputPath;
    private JComboBox<String> sheetComboBox;
    private JSpinner headerRowsSpinner;
    private JTextArea logArea;
    private JButton runButton;
    private JTable previewTable;

    public SimpleNormalizerDialog(Frame owner) {
        super(owner, "Simple Normalizer Tool", true);
        setSize(800, 600);
        setLocationRelativeTo(owner);
        setLayout(new BorderLayout(10, 10));

        JPanel mainPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 1.0;

        // Input File
        gbc.gridx = 0; gbc.gridy = 0; gbc.weightx = 0;
        mainPanel.add(new JLabel("Input Excel File:"), gbc);
        inputPath = new JTextField(30);
        inputPath.setEditable(false);
        gbc.gridx = 1; gbc.weightx = 1; mainPanel.add(inputPath, gbc);
        JButton chooseInputButton = new JButton("Choose...");
        chooseInputButton.addActionListener(e -> chooseInputFile());
        gbc.gridx = 2; gbc.weightx = 0; mainPanel.add(chooseInputButton, gbc);

        // Header Row Selection
        gbc.gridx = 0; gbc.gridy = 1;
        mainPanel.add(new JLabel("Header Rows:"), gbc);
        headerRowsSpinner = new JSpinner(new SpinnerNumberModel(1, 1, 100, 1));
        gbc.gridx = 1; gbc.gridwidth = 2; mainPanel.add(headerRowsSpinner, gbc);
        gbc.gridwidth = 1;

        // Sheet selection
        gbc.gridx = 0; gbc.gridy = 2;
        mainPanel.add(new JLabel("Sheet to Normalize:"), gbc);
        sheetComboBox = new JComboBox<>();
        sheetComboBox.setEnabled(false);
        sheetComboBox.addActionListener(e -> loadPreviewData());
        headerRowsSpinner.addChangeListener(e -> loadPreviewData());
        gbc.gridx = 1; gbc.gridwidth = 2; mainPanel.add(sheetComboBox, gbc);
        gbc.gridwidth = 1;

        add(mainPanel, BorderLayout.NORTH);

        previewTable = new JTable();
        JScrollPane previewScrollPane = new JScrollPane(previewTable);
        previewScrollPane.setBorder(BorderFactory.createTitledBorder("Preview (First 10 Rows)"));

        logArea = new JTextArea();
        logArea.setEditable(false);
        logArea.setFont(new Font("Monospaced", Font.PLAIN, 12));
        JScrollPane logScrollPane = new JScrollPane(logArea);
        logScrollPane.setBorder(BorderFactory.createTitledBorder("Log"));

        JSplitPane centerSplitPane = new JSplitPane(JSplitPane.VERTICAL_SPLIT, previewScrollPane, logScrollPane);
        centerSplitPane.setResizeWeight(0.6);
        add(centerSplitPane, BorderLayout.CENTER);

        JPanel actionPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        runButton = new JButton("Run Normalization");
        runButton.addActionListener(e -> runNormalization());
        actionPanel.add(runButton);
        add(actionPanel, BorderLayout.SOUTH);
    }

    private void chooseInputFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Select Input Excel File");
        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            inputFile = fileChooser.getSelectedFile();
            inputPath.setText(inputFile.getAbsolutePath());
            loadSheets();
        }
    }

    private void loadSheets() {
        if (inputFile == null) return;
        sheetComboBox.setEnabled(false);
        sheetComboBox.setModel(new DefaultComboBoxModel<>(new String[]{"Loading..."}));

        SwingWorker<List<String>, Void> worker = new SwingWorker<>() {
            @Override
            protected List<String> doInBackground() throws Exception {
                return ExcelUtil.getSheetNames(inputFile);
            }

            @Override
            protected void done() {
                try {
                    List<String> sheetNames = get();
                    sheetComboBox.setModel(new DefaultComboBoxModel<>(sheetNames.toArray(new String[0])));
                    sheetComboBox.setEnabled(true);
                    if (!sheetNames.isEmpty()) {
                        loadPreviewData();
                    }
                } catch (Exception e) {
                    sheetComboBox.setModel(new DefaultComboBoxModel<>(new String[]{"Error loading sheets"}));
                    JOptionPane.showMessageDialog(SimpleNormalizerDialog.this,
                            "Could not load sheets: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
        };
        worker.execute();
    }

    private void loadPreviewData() {
        if (inputFile == null || sheetComboBox.getSelectedItem() == null) return;

        String selectedSheet = (String) sheetComboBox.getSelectedItem();
        int headerRows = (int) headerRowsSpinner.getValue();

        SwingWorker<ExcelUtil.ExcelData, Void> worker = new SwingWorker<>() {
            @Override
            protected ExcelUtil.ExcelData doInBackground() throws Exception {
                return ExcelUtil.readExcel(inputFile, selectedSheet, headerRows, 10);
            }

            @Override
            protected void done() {
                try {
                    ExcelUtil.ExcelData excelData = get();
                    updatePreviewTable(excelData);
                } catch (Exception e) {
                    // Log error but don't show a popup, as it could be annoying
                    log("Error loading preview: " + e.getMessage());
                }
            }
        };
        worker.execute();
    }

    private void updatePreviewTable(ExcelUtil.ExcelData excelData) {
        if (excelData == null) {
            previewTable.setModel(new javax.swing.table.DefaultTableModel());
            return;
        }
        javax.swing.table.DefaultTableModel model = new javax.swing.table.DefaultTableModel(
            excelData.getData().stream().map(List::toArray).toArray(Object[][]::new),
            excelData.getHeaders().toArray()
        );
        previewTable.setModel(model);
    }

    private void runNormalization() {
        if (inputFile == null || sheetComboBox.getSelectedItem() == null) {
            JOptionPane.showMessageDialog(this, "Please select an input file and a sheet.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        String outputName = "normalized_" + inputFile.getName();
        File outputFile = new File(inputFile.getParent(), outputName);

        String selectedSheet = (String) sheetComboBox.getSelectedItem();
        int headerRows = (int) headerRowsSpinner.getValue();
        logArea.setText("");
        log("Starting normalization for sheet: " + selectedSheet);
        log("Output will be saved to: " + outputFile.getAbsolutePath());
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        runButton.setEnabled(false);

        new SwingWorker<Void, String>() {
            @Override
            protected Void doInBackground() throws Exception {
                StreamingExcelProcessor processor = new StreamingExcelProcessor();
                processor.processFile(inputFile, outputFile, selectedSheet, headerRows, this::publish);
                return null;
            }

            @Override
            protected void process(List<String> chunks) {
                for (String msg : chunks) {
                    log(msg);
                }
            }

            @Override
            protected void done() {
                try {
                    get(); // To catch exceptions
                    log("Normalization complete!");
                    String message = "Normalization complete!\nFile saved at: " + outputFile.getAbsolutePath();
                    JTextArea textArea = new JTextArea(message);
                    textArea.setEditable(false);
                    JScrollPane scrollPane = new JScrollPane(textArea);
                    scrollPane.setPreferredSize(new Dimension(400, 100));
                    JOptionPane.showMessageDialog(SimpleNormalizerDialog.this, scrollPane, "Success", JOptionPane.INFORMATION_MESSAGE);
                } catch (Exception e) {
                    String errorMsg = (e.getCause() != null) ? e.getCause().getMessage() : e.getMessage();
                    log("An error occurred: " + errorMsg);
                    JOptionPane.showMessageDialog(SimpleNormalizerDialog.this, "An error occurred:\n" + errorMsg, "Error", JOptionPane.ERROR_MESSAGE);
                } finally {
                    setCursor(Cursor.getDefaultCursor());
                    runButton.setEnabled(true);
                }
            }
        }.execute();
    }

    private void log(String message) {
        logArea.append(message + "\n");
    }
}
