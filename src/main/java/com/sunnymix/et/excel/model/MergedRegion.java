package com.sunnymix.et.excel.model;

import lombok.Getter;
import org.apache.poi.ss.util.CellRangeAddress;

@Getter
public class MergedRegion {

    private CellRangeAddress address;

    private int rowSpan;
    private int colSpan;

    private MergedRegion() {
    }

    public static MergedRegion of(CellRangeAddress address) {
        MergedRegion region = new MergedRegion();
        region.address = address;
        region.rowSpan = address.getLastRow() - address.getFirstRow() + 1;
        region.colSpan = address.getLastColumn() - address.getFirstColumn() + 1;
        return region;
    }

    public boolean startCell(Cell cell) {
        boolean firstRow = cell.getRow().getRawRowIdx() == address.getFirstRow();
        boolean firstCol = cell.getRawColIdx() == address.getFirstColumn();
        return firstRow && firstCol;
    }

    public boolean coverCell(Cell cell) {
        boolean cover = false;

        int rowIdx = cell.getRow().getRawRowIdx();
        int colIdx = cell.getRawColIdx();

        int firstRow = address.getFirstRow();
        int lastRow = address.getLastRow();
        int firstCol = address.getFirstColumn();
        int lastCol = address.getLastColumn();

        boolean sameRow = rowIdx == firstRow;
        boolean belowRow = rowIdx > firstRow
                && rowIdx <= lastRow;

        boolean amongCols = colIdx >= firstCol
                && colIdx <= lastCol;
        boolean amongColsWithoutFirst = amongCols
                && colIdx > firstCol;

        if (sameRow) {
            if (amongColsWithoutFirst) {
                cover = true;
            }
        } else if (belowRow) {
            if (amongCols) {
                cover = true;
            }
        }

        return cover;
    }

}
