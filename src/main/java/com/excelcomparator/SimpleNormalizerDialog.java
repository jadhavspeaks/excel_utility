package com.excelcomparator;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.util.List;

public class SimpleNormalizerDialog extends JDialog {

    private File inputFile;
    private JTextField inputPath;
    private JComboBox<String> sheetComboBox;
    private JTextArea logArea;
    private JButton runButton;

    public SimpleNormalizerDialog(Frame owner) {
        super(owner, "Simple Normalizer Tool", true);
        setSize(600, 400);
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

        // Sheet selection
        gbc.gridx = 0; gbc.gridy = 1;
        mainPanel.add(new JLabel("Sheet to Normalize:"), gbc);
        sheetComboBox = new JComboBox<>();
        sheetComboBox.setEnabled(false);
        gbc.gridx = 1; gbc.gridwidth = 2; mainPanel.add(sheetComboBox, gbc);
        gbc.gridwidth = 1;

        add(mainPanel, BorderLayout.NORTH);

        logArea = new JTextArea();
        logArea.setEditable(false);
        logArea.setFont(new Font("Monospaced", Font.PLAIN, 12));
        add(new JScrollPane(logArea), BorderLayout.CENTER);

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
                } catch (Exception e) {
                    sheetComboBox.setModel(new DefaultComboBoxModel<>(new String[]{"Error loading sheets"}));
                    JOptionPane.showMessageDialog(SimpleNormalizerDialog.this,
                            "Could not load sheets: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
        };
        worker.execute();
    }

    private void runNormalization() {
        if (inputFile == null || sheetComboBox.getSelectedItem() == null) {
            JOptionPane.showMessageDialog(this, "Please select an input file and a sheet.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Specify Output Excel File");
        fileChooser.setSelectedFile(new File(inputFile.getParent(), "normalized_" + inputFile.getName()));
        if (fileChooser.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) {
            return;
        }
        File outputFile = fileChooser.getSelectedFile();
         if (!outputFile.getName().toLowerCase().endsWith(".xlsx")) {
            outputFile = new File(outputFile.getParentFile(), outputFile.getName() + ".xlsx");
        }

        String selectedSheet = (String) sheetComboBox.getSelectedItem();
        logArea.setText("");
        log("Starting normalization for sheet: " + selectedSheet);
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        runButton.setEnabled(false);

        File finalOutputFile = outputFile;
        new SwingWorker<Void, String>() {
            @Override
            protected Void doInBackground() throws Exception {
                StreamingExcelProcessor processor = new StreamingExcelProcessor();
                processor.processFile(inputFile, finalOutputFile, selectedSheet, this::publish);
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
                    JOptionPane.showMessageDialog(SimpleNormalizerDialog.this, "Normalization complete!", "Success", JOptionPane.INFORMATION_MESSAGE);
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
