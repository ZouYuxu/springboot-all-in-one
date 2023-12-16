package com.example.jparest.utils;

import org.apache.poi.ss.usermodel.*;

public class ExcelUtil {

    public static void copyCell(Cell sourceCell, Cell targetCell) {
        if (sourceCell == null || targetCell == null) {
            return;
        }

        // 复制单元格样式
        CellStyle newCellStyle = targetCell.getSheet().getWorkbook().createCellStyle();
        newCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
        targetCell.setCellStyle(newCellStyle);

        // 根据单元格类型复制数据
        switch (sourceCell.getCellType()) {
            case BLANK:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case ERROR:
                targetCell.setCellValue(sourceCell.getErrorCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            case NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case STRING:
                targetCell.setCellValue(sourceCell.getRichStringCellValue());
                break;
            default:
                break;
        }
    }

    public static void copyColumn(Sheet sheet, int sourceColumnIndex, int targetColumnIndex) {
        int lastRowNum = sheet.getLastRowNum();

        for (int rowIndex = 0; rowIndex <= lastRowNum; rowIndex++) {
            Row row = sheet.getRow(rowIndex);

            // 判断目标列是否已有数据
            int insertIndex = targetColumnIndex;
            if (row.getCell(targetColumnIndex) != null) {
                // 如果已有数据，则在已有数据的前面插入
                insertIndex = row.getLastCellNum();
            }

            // 创建新单元格
            Cell targetCell = row.createCell(insertIndex);

            // 获取源单元格
            Cell sourceCell = (row.getCell(sourceColumnIndex) != null) ? row.getCell(sourceColumnIndex) : null;

            // 复制单元格
            copyCell(sourceCell, targetCell);
        }
    }

    public static void insertColumn(Sheet sheet, int columnIndex) {
        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);

            // 创建新单元格
            Cell newCell = row.createCell(columnIndex);

            // 设置新单元格的样式等
            // ...
        }
    }
}

