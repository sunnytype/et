package com.sunnymix.et.excel.model;

import lombok.Getter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

@Getter
public class CellStyle {

    private String fontColor = "111111";
    private String fontWeight = "normal";
    private String fillColor = "FFFFFF";

    private Border borderLeft;
    private Border borderRight;
    private Border borderTop;
    private Border borderBottom;

    private String textAlign = "left";
    private String verticalAlign = "top";

    public boolean isBlank() {
        return (borderLeft == null || borderLeft.isBlank())
                && (borderRight == null || borderRight.isBlank())
                && (borderTop == null || borderTop.isBlank())
                && (borderBottom == null || borderBottom.isBlank());
    }

    private CellStyle() {
    }

    public static CellStyle of(XSSFCellStyle rawStyle) {
        CellStyle style = new CellStyle();

        _extractColors(rawStyle, style);

        _extractAligns(rawStyle, style);

        _extractBorders(rawStyle, style);

        return style;
    }

    private static void _extractColors(XSSFCellStyle rawStyle, CellStyle style) {
        if (rawStyle.getFont() != null) {
            if (rawStyle.getFont().getXSSFColor() != null) {
                style.fontColor = rawStyle.getFont().getXSSFColor().getARGBHex().substring(2);
            }

            if (rawStyle.getFont().getBold()) {
                style.fontWeight = "bold";
            }
        }

        if (rawStyle.getFillForegroundColorColor() != null) {
            style.fillColor = rawStyle.getFillForegroundColorColor().getARGBHex().substring(2);
        }
    }

    private static void _extractBorders(XSSFCellStyle rawStyle, CellStyle style) {
        style.borderLeft = Border.of(rawStyle.getBorderLeftEnum(), rawStyle.getLeftBorderXSSFColor());
        style.borderRight = Border.of(rawStyle.getBorderRightEnum(), rawStyle.getRightBorderXSSFColor());
        style.borderTop = Border.of(rawStyle.getBorderTopEnum(), rawStyle.getTopBorderXSSFColor());
        style.borderBottom = Border.of(rawStyle.getBorderBottomEnum(), rawStyle.getBottomBorderXSSFColor());
    }

    private static void _extractAligns(XSSFCellStyle rawStyle, CellStyle style) {
        if (rawStyle.getAlignmentEnum() != null) {
            if (rawStyle.getAlignmentEnum().equals(HorizontalAlignment.CENTER)) {
                style.textAlign = "center";
            }
        }

        if (rawStyle.getVerticalAlignmentEnum() != null) {
            if (rawStyle.getVerticalAlignmentEnum().equals(VerticalAlignment.CENTER)) {
                style.verticalAlign = "middle";
            }
        }
    }

}
