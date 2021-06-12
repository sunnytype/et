package com.sunnymix.et.excel.model;

import lombok.Getter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

@Getter
public class Border {

    private int width;

    private String style = "none";

    private String color = "111111";

    private String css = "0 none";

    public boolean isBlank() {
        boolean widthIsBlank = width == 0;
        boolean styleIsBlank = style == null || style.equals("none");
        return widthIsBlank || styleIsBlank;
    }

    private Border() {
    }

    public static Border of(BorderStyle rawStyle, XSSFColor rawColor) {
        Border border = new Border();

        if (!rawStyle.equals(BorderStyle.NONE)) {
            if (rawStyle.equals(BorderStyle.DOUBLE)) {
                border.width = 3;
            } else if (rawStyle.equals(BorderStyle.MEDIUM)) {
                border.width = 2;
            } else {
                border.width = 1;
            }

            if (rawStyle.equals(BorderStyle.DASHED)) {
                border.style = "dashed";
            } else if (rawStyle.equals(BorderStyle.DOUBLE)) {
                border.style = "double";
            } else {
                border.style = "solid";
            }
        }

        if (rawColor != null && rawColor.getARGBHex() != null) {
            border.color = rawColor.getARGBHex().substring(2);
        }

        border.css = String.format("%spx %s #%s", border.width, border.style, border.color);
        return border;
    }

}
