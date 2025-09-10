package com.excelcomparator.ui;

import com.excelcomparator.model.ComparisonResult;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import java.awt.*;

public class StatusCellRenderer extends DefaultTableCellRenderer {
    private static final Color MATCH_COLOR = new Color(204, 255, 204); // Light Green
    private static final Color MISMATCH_COLOR = new Color(255, 255, 204); // Light Yellow
    private static final Color MISSING_COLOR = new Color(255, 204, 204); // Light Red

    @Override
    public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
        Component c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

        if (table.getModel().getColumnCount() > 2) {
            // Find the "Status" column index
            int statusColumnIndex = -1;
            for (int i = 0; i < table.getColumnCount(); i++) {
                if ("Status".equalsIgnoreCase(table.getColumnName(i))) {
                    statusColumnIndex = i;
                    break;
                }
            }

            if (statusColumnIndex != -1) {
                Object statusValue = table.getValueAt(row, statusColumnIndex);
                if (statusValue instanceof ComparisonResult.RowStatus) {
                    ComparisonResult.RowStatus status = (ComparisonResult.RowStatus) statusValue;
                    switch (status) {
                        case MATCH:
                            c.setBackground(MATCH_COLOR);
                            break;
                        case MISMATCH:
                            c.setBackground(MISMATCH_COLOR);
                            break;
                        case MISSING_IN_FILE_1:
                        case MISSING_IN_FILE_2:
                            c.setBackground(MISSING_COLOR);
                            break;
                        default:
                            c.setBackground(table.getBackground());
                            break;
                    }
                } else {
                     c.setBackground(table.getBackground());
                }
            }
        }

        if (isSelected) {
            c.setBackground(table.getSelectionBackground());
        }

        return c;
    }
}
