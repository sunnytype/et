package com.sunnymix.et.excel.model;

import lombok.Getter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

@Getter
public class Book {

    private String name;
    private List<Sheet> sheets;
    private int sheetsSize;

    private Book() {
    }

    public static Book of(File rawFile) {
        Book book = new Book();
        book.name = rawFile.getName();
        extractFile(rawFile, book);
        return book;
    }

    private static void extractFile(File file, Book book) {
        try {
            XSSFWorkbook rawBook = new XSSFWorkbook(file);

            int sheetsSize = rawBook.getNumberOfSheets();
            book.sheetsSize = sheetsSize;

            List<Sheet> sheets = new ArrayList<>();
            for (int i = 0; i < sheetsSize; i++) {
                XSSFSheet rawSheet = rawBook.getSheetAt(i);
                Sheet sheet = Sheet.of(rawSheet);
                sheets.add(sheet);
            }
            book.sheets = sheets;

        } catch (Throwable error) {
            System.out.println(error.getMessage());
        }
    }

}
