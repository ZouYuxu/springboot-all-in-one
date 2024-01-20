package com.example.jparest.utils;

import cn.hutool.core.util.NumberUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.awt.Color;
import java.util.Iterator;

@Slf4j
public class ExcelUtils<T> {

    public static final String TEMPLATE = "template";
    public static final String JSON = "json";


    private Class<T> eClass;

    public ExcelUtils(Class<T> eClass) {
        this.eClass = eClass;
    }

    public Class<T> getEClass() {
        return this.eClass;
    }

    // 获取行
    public static XSSFRow getRow(int i, XSSFSheet sheet, int templateRow) {
        XSSFRow row = sheet.getRow(i);

        if (row == null) {
            row = sheet.createRow(i);
            copyRow(sheet.getRow(templateRow), row);
        } else {
            row.setHeight(row.getHeight());
        }

        return row;
    }


    // copy样式
    public static void copyRow(Row srcRow, Row desRow) {
        desRow.setHeight(srcRow.getHeight());

        Iterator<Cell> it = srcRow.cellIterator();
        while (it.hasNext()) {
            Cell srcCell = it.next();

            srcCell.getCellStyle().setWrapText(true);
            Cell desCell = desRow.createCell(srcCell.getColumnIndex());
            desCell.setCellStyle(srcCell.getCellStyle());
            if (srcCell.getCellType().equals(CellType.FORMULA)) {
                desCell.setCellFormula(srcCell.getCellFormula());
            }
        }
    }


    public static String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    public static void copyCell(Cell sourceCell, Cell targetCell, String value) {
        if (sourceCell == null || targetCell == null) {
            return;
        }
        // 复制单元格样式
        CellStyle newCellStyle = targetCell.getSheet().getWorkbook().createCellStyle();
        newCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
        targetCell.setCellStyle(newCellStyle);
        targetCell.setCellValue(value);

        if (NumberUtil.isDouble(value)) {
            targetCell.setCellValue(NumberUtil.parseDouble(value));
        } else if (NumberUtil.isInteger(value)) {
            targetCell.setCellValue(NumberUtil.parseInt(value));
        } else {
            targetCell.setCellValue(value);
        }
    }

    public static void setCellStringValue(Cell cell, String value) {
        if (value != null) {
            cell.setCellValue(value);
        }
    }


    public static void setCellNumber(Cell cell, Integer value) {
        if (value != null) {
            cell.setCellValue(value);
        }
    }

    public static XSSFCellStyle getHeaderStyle(XSSFWorkbook wb) {
        XSSFCellStyle headerStyle = wb.createCellStyle();
        XSSFFont font = wb.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        // 设置字体为"微软雅黑"
        font.setFontName("微软雅黑");
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setFont(font);
        headerStyle.setFillForegroundColor(new XSSFColor(new Color(41, 129, 3), new DefaultIndexedColorMap())); // 设置背景颜色为 #298103
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        return headerStyle;
    }

    public static XSSFCellStyle getCenterStyle(XSSFWorkbook wb) {
        XSSFCellStyle center = wb.createCellStyle();
        XSSFFont font = wb.createFont();
        font.setFontName("微软雅黑");
        center.setFont(font);
        center.setBorderTop(BorderStyle.THIN);
        center.setBorderBottom(BorderStyle.THIN);
        center.setBorderLeft(BorderStyle.THIN);
        center.setBorderRight(BorderStyle.THIN);
        center.setAlignment(HorizontalAlignment.CENTER);
        return center;
    }

    public static XSSFCellStyle getDataLeftCenterStyle(XSSFWorkbook wb) {
        XSSFCellStyle center = wb.createCellStyle();
        XSSFFont font = wb.createFont();
        font.setFontName("微软雅黑");
        center.setFont(font);
        center.setBorderTop(BorderStyle.THIN);
        center.setBorderBottom(BorderStyle.THIN);
        center.setBorderLeft(BorderStyle.THIN);
        center.setBorderRight(BorderStyle.THIN);
        center.setAlignment(HorizontalAlignment.LEFT);
        return center;
    }

    public static XSSFCellStyle getPicCenterStyle(XSSFWorkbook wb) {
        XSSFCellStyle center = wb.createCellStyle();
        XSSFFont font = wb.createFont();
        font.setFontName("微软雅黑");
        center.setFont(font);
        center.setBorderTop(BorderStyle.THIN);
        center.setBorderBottom(BorderStyle.THIN);
        center.setBorderLeft(BorderStyle.THIN);
        center.setBorderRight(BorderStyle.THIN);
        center.setAlignment(HorizontalAlignment.LEFT);
        return center;
    }

    public static XSSFCellStyle getHighLightStyle(XSSFWorkbook wb, Boolean isHighLight) {
        XSSFCellStyle highLightStyle = wb.createCellStyle();
        XSSFFont font = wb.createFont();
        font.setColor(IndexedColors.BLACK.getIndex());
        // 设置字体为"微软雅黑"
        font.setFontName("微软雅黑");
        highLightStyle.setFont(font);
        if (isHighLight) {
            highLightStyle.setFillForegroundColor(new XSSFColor(new Color(255, 255, 153), new DefaultIndexedColorMap())); // 设置背景颜色为 #298103
        } else {
            highLightStyle.setFillForegroundColor(new XSSFColor(new Color(226, 239, 218), new DefaultIndexedColorMap())); // 设置背景颜色为 #298103
        }
        highLightStyle.setBorderTop(BorderStyle.THIN);
        highLightStyle.setBorderBottom(BorderStyle.THIN);
        highLightStyle.setBorderLeft(BorderStyle.THIN);
        highLightStyle.setBorderRight(BorderStyle.THIN);
        highLightStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return highLightStyle;
    }


}