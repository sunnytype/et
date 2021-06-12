package com.sunnymix.et.excel.model;

import lombok.Getter;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRelation;

import java.text.SimpleDateFormat;
import java.util.Date;

@Getter
public class Cell {

    private Row row;

    private int rawColIdx;
    private int colIdx;

    private CellType type;

    private String rawValue;
    private String value;

    private boolean merged;

    private int rowSpan = 1;
    private int colSpan = 1;

    private boolean cover;

    private boolean link;
    private String linkLocation;

    private CellStyle style;

    public boolean isNotCover() {
        return !cover;
    }

    public boolean isBlank() {
        boolean valueIsBlank = value == null || value.isBlank();
        boolean styleIsBlank = style == null || style.isBlank();
        return valueIsBlank && styleIsBlank;
    }

    public boolean isNotBlank() {
        return !isBlank();
    }

    private Cell() {
    }

    public static Cell of(Row row, int rawColIdx, int colIdx, XSSFCell rawCell) {
        Cell cell = new Cell();
        cell.row = row;

        cell.rawColIdx = rawColIdx;
        cell.colIdx = colIdx;

        cell.type = CellType.of(rawCell.getCellType());

        cell.rawValue = extractValue(rawCell, cell.type);
        cell.value = cell.rawValue.replaceAll("\n", "<br>");

        cell.style = CellStyle.of(rawCell.getCellStyle());

        _extractLink(rawCell, cell);

        _mergedRegion(row, cell);

        return cell;
    }

    private static String extractValue(XSSFCell rawCell, CellType cellType) {
        String value = "?";

        if (cellType.isBlank()) {
            value = "";
        } else if (cellType.isString()) {
            value = rawCell.getStringCellValue();
        } else if (cellType.isBool()) {
            value = rawCell.getBooleanCellValue() ? "true" : "false";
        } else if (cellType.isNumeric()) {
            value = extractNum(rawCell);
        } else {
            value = "??";
        }

        return value;
    }

    private static String extractNum(XSSFCell rawCell) {
        String value = "?";
        try {
            Date dateCellValue = rawCell.getDateCellValue();
            value = new SimpleDateFormat("M/d").format(dateCellValue);
        } catch (NumberFormatException notDateNumError) {
            value = String.format("%s", rawCell.getNumericCellValue());
        } catch (Throwable error) {
            value = "??";
        }
        return value;
    }

    private static void _extractLink(XSSFCell rawCell, Cell cell) {
        XSSFHyperlink rawLink = rawCell.getHyperlink();
        String link = null;
        if (rawLink != null) {
            try {
                PackagePart packagePart = rawCell.getRow().getSheet().getPackagePart();
                PackageRelationshipCollection hyperRels = packagePart.getRelationshipsByType(XSSFRelation.SHEET_HYPERLINKS.getRelation());
                PackageRelationship relationshipByID = hyperRels.getRelationshipByID(rawLink.getCTHyperlink().getId());
                link = relationshipByID.getTargetURI().toString();
                if (rawLink.getCTHyperlink().getLocation() != null) {
                    link = String.format("%s#%s", link, rawLink.getCTHyperlink().getLocation());
                }
            } catch (Throwable error) {
            }
        }

        if (link != null) {
            cell.link = true;
            cell.linkLocation = link;
        }

    }

    private static void _mergedRegion(Row row, Cell cell) {
        for (MergedRegion region : row.getSheet().getMergedRegions()) {
            if (region.startCell(cell)) {
                cell.merged = true;
                cell.rowSpan = region.getRowSpan();
                cell.colSpan = region.getColSpan();
                break;
            } else if (region.coverCell(cell)) {
                cell.cover = true;
                break;
            }
        }
    }

}
