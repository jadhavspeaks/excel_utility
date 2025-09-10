package com.excelcomparator;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class SimpleFilterDialog extends JDialog {

    private File inputFile;
    private JTextField inputFilePath;
    private JComboBox<String> sheetCombo;
    private JComboBox<String> keyColumnCombo;
    private JTextArea keywordsTextArea;
    private JTextArea logArea;
    private JButton startButton;

    public SimpleFilterDialog(Frame owner) {
        super(owner, "Simple Keyword Filter", true);
        setSize(600, 700);
        setLocationRelativeTo(owner);
        setLayout(new BorderLayout(10, 10));

        add(createConfigPanel(), BorderLayout.NORTH);
        add(createLogPanel(), BorderLayout.CENTER);
        add(createButtonPanel(), BorderLayout.SOUTH);
    }

    private JPanel createConfigPanel() {
        JPanel panel = new JPanel();
        panel.setLayout(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder("Configuration"));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        // Input File
        gbc.gridx = 0; gbc.gridy = 0;
        panel.add(new JLabel("Input Excel File:"), gbc);
        inputFilePath = new JTextField(30);
        inputFilePath.setEditable(false);
        gbc.gridx = 1; gbc.gridwidth = 2; panel.add(inputFilePath, gbc);
        JButton chooseFileButton = new JButton("Choose...");
        chooseFileButton.addActionListener(e -> chooseInputFile());
        gbc.gridx = 3; gbc.gridwidth = 1; panel.add(chooseFileButton, gbc);

        // Sheet Selection
        gbc.gridx = 0; gbc.gridy = 1;
        panel.add(new JLabel("Sheet to Filter:"), gbc);
        sheetCombo = new JComboBox<>();
        sheetCombo.addActionListener(e -> loadColumnsForSheet());
        gbc.gridx = 1; gbc.gridwidth = 3; panel.add(sheetCombo, gbc);

        // Key Column
        gbc.gridx = 0; gbc.gridy = 2;
        panel.add(new JLabel("Key Column:"), gbc);
        keyColumnCombo = new JComboBox<>();
        gbc.gridx = 1; gbc.gridwidth = 3; panel.add(keyColumnCombo, gbc);

        // Keywords
        gbc.gridx = 0; gbc.gridy = 3; gbc.anchor = GridBagConstraints.NORTHWEST;
        panel.add(new JLabel("Keywords (one per line):"), gbc);
        keywordsTextArea = new JTextArea(10, 30);
        gbc.gridx = 1; gbc.gridy = 3; gbc.gridwidth = 3; gbc.fill = GridBagConstraints.BOTH;
        panel.add(new JScrollPane(keywordsTextArea), gbc);


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
                List<String> sheetNames = new ArrayList<>();
                try (InputStream inputStream = new FileInputStream(inputFile);
                     Workbook workbook = new XSSFWorkbook(inputStream)) {
                    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                        sheetNames.add(workbook.getSheetName(i));
                    }
                }
                return sheetNames;
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

    private void loadColumnsForSheet() {
        String selectedSheet = (String) sheetCombo.getSelectedItem();
        if (inputFile == null || selectedSheet == null) return;

        SwingWorker<List<String>, Void> worker = new SwingWorker<>() {
            @Override
            protected List<String> doInBackground() throws Exception {
                return ExcelUtil.getHeaders(inputFile, selectedSheet);
            }

            @Override
            protected void done() {
                try {
                    List<String> headers = get();
                    keyColumnCombo.setModel(new DefaultComboBoxModel<>(headers.toArray(new String[0])));
                } catch (Exception e) {
                    log("Error loading columns: " + e.getMessage());
                }
            }
        };
        worker.execute();
    }

    private void startFiltering() {
        if (inputFile == null || sheetCombo.getSelectedItem() == null || keyColumnCombo.getSelectedItem() == null || keywordsTextArea.getText().isBlank()) {
            JOptionPane.showMessageDialog(this, "Please fill all fields.", "Configuration Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Save Filtered Output");
        fileChooser.setSelectedFile(new File("filtered_output.xlsx"));
        if (fileChooser.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) {
            return;
        }
        File outputFile = fileChooser.getSelectedFile();

        String sheetName = (String) sheetCombo.getSelectedItem();
        int keyColumnIndex = keyColumnCombo.getSelectedIndex();
        List<String> keywords = Arrays.asList(keywordsTextArea.getText().split("\\s*\\r?\\n\\s*"));

        log("Starting filter process...");
        startButton.setEnabled(false);

        SwingWorker<Void, String> worker = new SwingWorker<>() {
            @Override
            protected Void doInBackground() throws Exception {
                ExcelFilter filter = new ExcelFilter();
                filter.processFile(inputFile, outputFile, sheetName, keyColumnIndex, keywords, this::publish);
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
                } catch (Exception e) {
                    log("Error during filtering: " + e.getMessage());
                    e.printStackTrace();
                    JOptionPane.showMessageDialog(SimpleFilterDialog.this, "An error occurred: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                } finally {
                    startButton.setEnabled(true);
                }
            }
        };

        worker.execute();
    }

    private void log(String message) {
        logArea.append(message + "\n");
    }
}
