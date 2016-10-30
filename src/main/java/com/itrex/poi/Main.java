package com.itrex.poi;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

public class Main {

    public static void main(String[] args) throws IOException {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        ByteArrayOutputStream os = new ByteArrayOutputStream();

        try {

            createHeader(workbook, sheet);
            createContent(workbook, sheet);

            workbook.write(os);

            os.close();

        } catch (IOException e) {
            e.printStackTrace();
        }


        OutputStream is = new FileOutputStream("workbook.xlsx");

        is.write(os.toByteArray());

        is.close();
    }

    private static void createHeader(Workbook workbook, Sheet sheet) {
        Row row = sheet.createRow(0);
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();

        font.setBold(true);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFont(font);

        createCell(row, 0, "Column 1", style);
        createCell(row, 1, "Column 2", style);
        createCell(row, 2, "Column 3", style);
        createCell(row, 3, "Column 4", style);
    }

    private static void createContent(Workbook workbook, Sheet sheet) {
        CellStyle center = workbook.createCellStyle();
        center.setAlignment(CellStyle.ALIGN_CENTER);

        CellStyle left = workbook.createCellStyle();
        left.setAlignment(CellStyle.ALIGN_LEFT);

        CellStyle right = workbook.createCellStyle();
        right.setAlignment(CellStyle.ALIGN_RIGHT);

        for (int i = 0; i < 10; i++) {
            Row row = sheet.createRow(i + 1);

            createCell(row, 0, "Text", center);
            createCell(row, 1, "TextTextText", center);
            createCell(row, 2, "TextTextTextTextTextTextText", left);
            createCell(row, 3, "TextTextTextTextTextTextTextTextText", right);
        }

        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);
    }

    private static void createCell(Row row, int col, String text, CellStyle style) {
        Cell cell = row.createCell(col);

        cell.setCellValue(text);
        cell.setCellStyle(style);
    }

}
