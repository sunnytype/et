package com.sunnymix.et.excel.model;

import lombok.Getter;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.List;

@Getter
public class Sheet {

    private String name;

    private List<MergedRegion> mergedRegions;

    private List<Row> rows;
    private int rowsSize;

    private Sheet() {
    }

    public static Sheet of(XSSFSheet rawSheet) {
        Sheet sheet = new Sheet();
        _extractSheet(rawSheet, sheet);
        return sheet;
    }

    private static void _extractSheet(XSSFSheet rawSheet, Sheet sheet) {
        sheet.name = rawSheet.getSheetName();
        _extractMergedRegions(rawSheet, sheet);
        _extractRows(rawSheet, sheet);
    }

    private static void _extractMergedRegions(XSSFSheet rawSheet, Sheet sheet) {
        int regionsSize = rawSheet.getNumMergedRegions();

        List<MergedRegion> regions = new ArrayList<>();
        for (int i = 0; i < regionsSize; i++) {
            CellRangeAddress regionAddress = rawSheet.getMergedRegion(i);
            MergedRegion region = MergedRegion.of(regionAddress);
            regions.add(region);
        }
        sheet.mergedRegions = regions;
    }

    private static void _extractRows(XSSFSheet rawSheet, Sheet sheet) {
        int rowsSize = rawSheet.getPhysicalNumberOfRows();
        sheet.rowsSize = rowsSize;

        int firstRowIdx = rawSheet.getFirstRowNum();

        List<Row> rows = new ArrayList<>();
        for (int rowIdx = 0; rowIdx < rowsSize; rowIdx++) {
            int rawRowIdx = rowIdx + firstRowIdx;
            XSSFRow rawRow = rawSheet.getRow(rawRowIdx);
            Row row = Row.of(sheet, rawRowIdx, rowIdx, rawRow);
            if (row.isNotBlank()) {
                rows.add(row);
            }
        }
        sheet.rows = rows;
    }

}
