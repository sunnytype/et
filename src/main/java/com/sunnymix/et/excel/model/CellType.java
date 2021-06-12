package com.sunnymix.et.excel.model;

import lombok.Getter;

import static org.apache.poi.ss.usermodel.Cell.*;

@Getter
public class CellType {

    private int rawType;

    private boolean numeric;

    private boolean string;

    private boolean formula;

    private boolean blank;

    private boolean bool;

    private boolean error;

    private CellType() {
    }

    public static CellType of(int rawType) {
        CellType cellType = new CellType();

        cellType.rawType = rawType;

        cellType.numeric = rawType == CELL_TYPE_NUMERIC;
        cellType.string = rawType == CELL_TYPE_STRING;
        cellType.formula = rawType == CELL_TYPE_FORMULA;
        cellType.blank = rawType == CELL_TYPE_BLANK;
        cellType.bool = rawType == CELL_TYPE_BOOLEAN;
        cellType.error = rawType == CELL_TYPE_ERROR;

        return cellType;
    }

}
