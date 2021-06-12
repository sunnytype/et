package com.sunnymix.et.excel.model;

import lombok.Getter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.ss.usermodel.Row.RETURN_NULL_AND_BLANK;

@Getter
public class Row {

    private Sheet sheet;

    private int rawRowIdx;
    private int rowIdx;

    private List<Cell> cells;
    private int cellsSize;

    private boolean head;

    public boolean isBlank() {
        return cellsSize == 0;
    }

    public boolean isNotBlank() {
        return !isBlank();
    }

    private Row() {
    }

    public static Row of(Sheet sheet, int rawRowIdx, int rowIdx, XSSFRow rawRow) {
        Row row = new Row();

        row.sheet = sheet;

        row.rawRowIdx = rawRowIdx;
        row.rowIdx = rowIdx;

        row.head = rowIdx == 0;

        extractCells(rawRow, row);
        return row;
    }

    private static void extractCells(XSSFRow rawRow, Row row) {
        int cellsSize = rawRow.getPhysicalNumberOfCells();

        int firstColIdx = rawRow.getFirstCellNum();

        List<Cell> cells = new ArrayList<>();
        for (int colIdx = 0; colIdx < cellsSize; colIdx++) {
            int rawColIdx = colIdx + firstColIdx;
            XSSFCell rawCell = rawRow.getCell(rawColIdx, RETURN_NULL_AND_BLANK);
            if (rawCell != null) {
                Cell cell = Cell.of(row, rawColIdx, colIdx, rawCell);
                if (cell.isNotCover() && cell.isNotBlank()) {
                    cells.add(cell);
                }
            }
        }
        row.cells = cells;
        row.cellsSize = cells.size();
    }

}
