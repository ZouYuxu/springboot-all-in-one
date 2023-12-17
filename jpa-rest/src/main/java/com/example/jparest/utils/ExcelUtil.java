package com.example.jparest.utils;

import lombok.Value;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

public class ExcelUtil {

    public static void copyCell(Cell sourceCell, Cell targetCell, String value) {
        if (sourceCell == null || targetCell == null) {
            return;
        }

        // 复制单元格样式
        CellStyle newCellStyle = targetCell.getSheet().getWorkbook().createCellStyle();
        newCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
        targetCell.setCellStyle(newCellStyle);

        // 根据单元格类型复制数据
        switch (sourceCell.getCellType()) {
            case BLANK, STRING:
                targetCell.setCellValue(value != null ? value : sourceCell.getStringCellValue());
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
            default:
                break;
        }
    }

    public static void copyColumn(Sheet sheet, int sourceColumnIndex, int targetColumnIndex) {
        copyColumn(sheet, sourceColumnIndex, targetColumnIndex, -1, null);
    }

    public static void copyColumn(Sheet sourceSheet, Sheet targetSheet, int sourceColumnIndex, int targetColumnIndex) {
    }

    // copy column from one sheet to another
    public static void copyColumn(Sheet sourceSheet, Sheet targetSheet, int sourceColumnIndex, int targetColumnIndex, List<String> list) {
        for (int i = 0; i < sourceSheet.getPhysicalNumberOfRows(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            Row targetRow = targetSheet.getRow(i);
            String newValue = list.get(i);
            // fix bug：brand column missing
            if (targetRow == null) {
                targetRow = targetSheet.createRow(i);
            }
            if (sourceRow != null) {
                Cell sourceCell = sourceRow.getCell(sourceColumnIndex);
                Cell newCell = targetRow.createCell(targetColumnIndex);
                copyCell(sourceCell, newCell, newValue);
            }
        }

    }

    // todo text可以不用传递过来，之后再赋值也可以
    public static void copyColumn(Sheet sheet, int sourceColumnIndex, int targetColumnIndex, int textRowIndex, String newValue) {
        int lastRowNum = sheet.getLastRowNum();

        // 向右平移
        sheet.shiftColumns(targetColumnIndex, sheet.getRow(0).getLastCellNum(), 1);
        int newSourceColumnIndex = sourceColumnIndex + 1;
        for (int rowIndex = 0; rowIndex <= lastRowNum; rowIndex++) {
            Row row = sheet.getRow(rowIndex);

            // 判断目标列是否已有数据
            int insertIndex = targetColumnIndex;
//            if (row.getCell(targetColumnIndex) != null) {
//                // 如果已有数据，则在已有数据的前面插入
//                insertIndex = row.getLastCellNum();
//            }

            // 创建新单元格
            Cell targetCell = row.createCell(insertIndex);

            // 获取源单元格
            Cell sourceCell = (row.getCell(newSourceColumnIndex) != null) ? row.getCell(newSourceColumnIndex) : null;

            // 复制单元格
            copyCell(sourceCell, targetCell, textRowIndex == rowIndex ? newValue : null);
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

