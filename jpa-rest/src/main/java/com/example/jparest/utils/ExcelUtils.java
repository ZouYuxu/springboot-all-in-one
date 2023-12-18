package com.example.jparest.utils;

import com.fasterxml.jackson.databind.JsonNode;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.awt.Color;
import java.util.*;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class ExcelUtils<T> {

    public static final String TEMPLATE = "template";

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

    // copy column from one sheet to another
    public static void copyColumn(Sheet sourceSheet, Sheet targetSheet, int sourceColumnIndex, int targetColumnIndex, List<String > list, int dataListRow, LinkedList<JsonNode> node, JsonNode data, HashMap<String, Integer> varLineMap) {
        int i = 0;
        List<String> ss = List.of("我真的", "kusi");

        for (; i < dataListRow; i++) {
            boolean flag = i >= dataListRow;
            int index = i - dataListRow;
            Row sourceRow = sourceSheet.getRow(flag ? dataListRow : i);
            Row targetRow = targetSheet.getRow(i);
            String newValue = flag ? ss.get(index) : list.get(i);
            // fix bug：brand column missing
            if (targetRow == null) {
                targetRow = targetSheet.createRow(i);
            }
            if (sourceRow != null) {
                Cell sourceCell = sourceRow.getCell(sourceColumnIndex);
//                fun.call(sourceValue);
                Cell newCell = targetRow.createCell(targetColumnIndex);
                copyCell(sourceCell, newCell, newValue, dataListRow);
            }
        }


        String sourceValue = sourceSheet.getRow(dataListRow).getCell(sourceColumnIndex).getStringCellValue();
        sourceValue = sourceValue.substring(1, sourceValue.length() - 1);
        String newPath = Pattern.compile("\\[(.*?)\\]").matcher(sourceValue).replaceAll(m -> {
            String group = m.group(1);
            // todo 在前面解决好路径的问题
            group = group.replaceAll("\\.", "/");
            String[] split = group.split("/");
            String variable = split[0];
            String collect = "/"+Arrays.stream(split).skip(1).collect(Collectors.joining("/"));
//            collect =
            Integer integer = varLineMap.get(variable);
            JsonNode jsonNode = node.get(integer);
            String text;
            if (jsonNode.isObject()) {
                text = jsonNode.at(collect).asText();
            } else {
                text = jsonNode.asText();
            }
            System.out.println(text);
            return "/" + text;
        });

        // to:　data/sub/cost/point/10
        newPath = newPath.replaceAll("\\.", "/");
        // to: /sub/cost/point/10
        newPath = newPath.substring(newPath.indexOf("/"));
        List<String> strings = new ArrayList<>();
        if (data.isArray()) {
            for (JsonNode item : data) {
                strings.add(item.at(newPath).asText());
            }
        }
        for (; i < dataListRow + strings.size(); i++) {
            boolean flag = i >= dataListRow;
            int index = i - dataListRow;
            Row sourceRow = sourceSheet.getRow(flag ? dataListRow : i);
            Row targetRow = targetSheet.getRow(i);
            String newValue = flag ? strings.get(index) : list.get(i);
            // fix bug：brand column missing
            if (targetRow == null) {
                targetRow = targetSheet.createRow(i);
            }
            if (sourceRow != null) {
                Cell sourceCell = sourceRow.getCell(sourceColumnIndex);
//                fun.call(sourceValue);
                Cell newCell = targetRow.createCell(targetColumnIndex);
                copyCell(sourceCell, newCell, newValue, dataListRow);
            }
        }
    }

    public static void copyCell(Cell sourceCell, Cell targetCell, String value, int dataListRow) {
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