package com.excelcomparator.ui;

import javax.swing.*;
import java.awt.*;
import java.util.List;

public class FilterConditionRow extends JPanel {
    private JComboBox<String> columnCombo;
    private JComboBox<String> operatorCombo;
    private JTextField valueField;

    public FilterConditionRow(List<String> columnNames) {
        setLayout(new FlowLayout(FlowLayout.LEFT));

        columnCombo = new JComboBox<>(columnNames.toArray(new String[0]));
        operatorCombo = new JComboBox<>(new String[]{"EQUALS", "CONTAINS", "NOT_CONTAINS", "STARTS_WITH", "ENDS_WITH"});
        valueField = new JTextField(20);

        add(new JLabel("Column:"));
        add(columnCombo);
        add(new JLabel("Operator:"));
        add(operatorCombo);
        add(new JLabel("Value:"));
        add(valueField);
    }

    public String getSelectedColumn() {
        return (String) columnCombo.getSelectedItem();
    }

    public String getSelectedOperator() {
        return (String) operatorCombo.getSelectedItem();
    }

    public String getValue() {
        return valueField.getText();
    }

    public void setColumnNames(List<String> columnNames) {
        String selected = getSelectedColumn();
        columnCombo.setModel(new DefaultComboBoxModel<>(columnNames.toArray(new String[0])));
        columnCombo.setSelectedItem(selected);
    }
}
