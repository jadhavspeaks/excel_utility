package com.excelcomparator;

import javax.swing.*;
import java.awt.*;
import java.io.File;

public class SimpleNormalizerDialog extends JDialog {

    private File inputFile, outputFile;
    private JTextField inputPath, outputPath;
    private JTextArea logArea;

    public SimpleNormalizerDialog(Frame owner) {
        super(owner, "Simple Normalizer Tool", false);
        setSize(600, 400);
        setLocationRelativeTo(owner);
        setLayout(new BorderLayout(10, 10));

        JPanel mainPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        gbc.gridx = 0; gbc.gridy = 0;
        mainPanel.add(new JLabel("Input Excel File:"), gbc);
        inputPath = new JTextField(30);
        inputPath.setEditable(false);
        gbc.gridx = 1; mainPanel.add(inputPath, gbc);
        JButton chooseInputButton = new JButton("Choose...");
        chooseInputButton.addActionListener(e -> chooseFile(true));
        gbc.gridx = 2; mainPanel.add(chooseInputButton, gbc);

        gbc.gridx = 0; gbc.gridy = 1;
        mainPanel.add(new JLabel("Output Excel File:"), gbc);
        outputPath = new JTextField(30);
        outputPath.setEditable(false);
        gbc.gridx = 1; mainPanel.add(outputPath, gbc);
        JButton chooseOutputButton = new JButton("Choose...");
        chooseOutputButton.addActionListener(e -> chooseFile(false));
        gbc.gridx = 2; mainPanel.add(chooseOutputButton, gbc);

        add(mainPanel, BorderLayout.NORTH);

        logArea = new JTextArea();
        logArea.setEditable(false);
        logArea.setFont(new Font("Monospaced", Font.PLAIN, 12));
        add(new JScrollPane(logArea), BorderLayout.CENTER);

        JPanel actionPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        JButton runButton = new JButton("Run Normalization");
        runButton.addActionListener(e -> runNormalization());
        actionPanel.add(runButton);
        add(actionPanel, BorderLayout.SOUTH);
    }

    private void chooseFile(boolean isInput) {
        JFileChooser fileChooser = new JFileChooser();
        if (isInput) {
            fileChooser.setDialogTitle("Select Input Excel File");
            if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
                inputFile = fileChooser.getSelectedFile();
                inputPath.setText(inputFile.getAbsolutePath());
            }
        } else {
            fileChooser.setDialogTitle("Specify Output Excel File");
            if (fileChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
                outputFile = fileChooser.getSelectedFile();
                if (!outputFile.getName().toLowerCase().endsWith(".xlsx")) {
                    outputFile = new File(outputFile.getParentFile(), outputFile.getName() + ".xlsx");
                }
                outputPath.setText(outputFile.getAbsolutePath());
            }
        }
    }

    private void runNormalization() {
        if (inputFile == null || outputFile == null) {
            JOptionPane.showMessageDialog(this, "Please select both input and output files.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        logArea.setText("Processing... Please wait.");
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        new SwingWorker<String, Void>() {
            @Override
            protected String doInBackground() throws Exception {
                StreamingExcelProcessor processor = new StreamingExcelProcessor();
                return processor.normalize(inputFile, outputFile);
            }

            @Override
            protected void done() {
                try {
                    String log = get();
                    logArea.setText(log);
                    JOptionPane.showMessageDialog(SimpleNormalizerDialog.this, "Normalization complete!", "Success", JOptionPane.INFORMATION_MESSAGE);
                } catch (Exception e) {
                    logArea.setText("An error occurred:\n" + e.getCause().getMessage());
                    JOptionPane.showMessageDialog(SimpleNormalizerDialog.this, "An error occurred during normalization:\n" + e.getCause().getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                } finally {
                    setCursor(Cursor.getDefaultCursor());
                }
            }
        }.execute();
    }
}
