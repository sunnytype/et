package com.sunnymix.et.excel.model.view;

import com.sunnymix.et.excel.model.Book;
import com.sunnymix.et.excel.model.Cell;
import com.sunnymix.et.excel.model.Row;
import com.sunnymix.et.excel.model.Sheet;
import lombok.Getter;

@Getter
public class TextView {

    private String text;

    private TextView() {
    }

    public static TextView of(Book book) {
        TextView textView = new TextView();

        StringBuilder text = new StringBuilder("");

        for (int sheetIdx = 0; sheetIdx < book.getSheets().size(); sheetIdx++) {
            Sheet sheet = book.getSheets().get(sheetIdx);

            String sheetText = sheet.getName() + ": \n" +
                    sheetText(sheet) + "\n\n";
            text.append(sheetText);
        }

        textView.text = text.toString();
        return textView;
    }

    private static String sheetText(Sheet sheet) {
        StringBuilder sheetText = new StringBuilder();

        int rowSize = sheet.getRows().size();

        boolean tabRow = false;
        int tabRowIdx = -1;
        int tabRowMaxIdx = 0;
        boolean tabStart = false;
        boolean tabStartFollower = false;

        for (int rowIdx = 0; rowIdx < rowSize; rowIdx++) {
            Row row = sheet.getRows().get(rowIdx);

            if (rowIdx == 0) {
                continue;
            }

            StringBuilder rowText = new StringBuilder();

            for (int cellIdx = 0; cellIdx < row.getCells().size(); cellIdx++) {
                Cell cell = row.getCells().get(cellIdx);

                boolean firstCell = cellIdx == 0;

                if (firstCell) {
                    int rowSpan = cell.getRowSpan();

                    if (rowSpan > 1) {
                        tabRowMaxIdx = rowSpan - 1;
                        tabRow = true;
                        tabRowIdx = 0;
                    }
                }

                tabStart = tabRow && tabRowIdx == 0 && firstCell;

                String cellText = cell.getRawValue().replaceAll("\n", " ");

                if (cellIdx > 0 && !cellText.isBlank()) {
                    if (tabStartFollower) {
                        tabStartFollower = false;
                    } else {
                        cellText = ", " + cellText;
                    }
                }

                if (firstCell) {
                    if (tabStart) {
                        rowText.append("▶ ").append(cellText).append(": \n").append("    ▷ ");
                        tabStartFollower = true;
                    } else {
                        if (tabRow) {
                            rowText.append("    ▷ ");
                        } else {
                            rowText.append("▶ ");
                        }
                        rowText.append(cellText);
                    }
                } else {
                    rowText.append(cellText);
                }

            } // End of cells

            String rowTextRaw = rowText.toString();

            if (!rowTextRaw.isBlank() && !tabStart) {
                rowTextRaw += "\n";
            }

            sheetText.append(rowTextRaw);

            tabRowIdx += 1;

            if (tabRowIdx > tabRowMaxIdx) {
                tabRow = false;
                tabStart = false;
                tabRowIdx = -1;
            }

        } // End of rows

        return sheetText.toString();
    }

}
