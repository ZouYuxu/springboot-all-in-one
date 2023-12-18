package com.example.jparest.utils;

import cn.hutool.core.io.FileUtil;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;

import java.io.*;
import java.util.*;


@Slf4j
public class EvaluationAnalyseDownloadService {
    private int ind = 0;
    private int fastIndex = 0;

    public void downloadEvaluationAnalyse() throws Exception {

        String step = "[1.0].Read template";
        byte[] bytes = FileUtil.readBytes(ExcelUtils.TEMPLATE + File.separator + "easy.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes));
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            XSSFSheet sheet = wb.getSheet("easy");
            for (int rowIndex = 0; rowIndex < sheet.getLastRowNum(); rowIndex++) {
                XSSFRow row = sheet.getRow(rowIndex);
                if (row != null) {
                    // 遍历每个单元格
                    for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
                        Cell cell = row.getCell(cellIndex);

                        if (cell != null) {
                            // 输出单元格内容
                            System.out.println("[" + rowIndex + "," + cellIndex + "]: " + getCellValue(cell));
                        }
                    }
                }

            }

        } catch (Exception e) {
            throw new Exception(step + "生成Excel发生错误！");
        }
    }

    public void template() throws Exception {
        ObjectMapper mapper = new ObjectMapper();
        InputStream inputStream = new ClassPathResource("json/var.json").getInputStream();
        InputStream dataInputStream = new ClassPathResource("json/data.json").getInputStream();
        JsonNode var = mapper.readTree(inputStream);
        JsonNode data = mapper.readTree(dataInputStream);
        String step = "[1.0].Read template";
        byte[] bytes = FileUtil.readBytes("template" + File.separator + "easy.xlsx");
        String destFilePath = new ClassPathResource("template/dest.xlsx").getFile().getAbsolutePath();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes));
             XSSFWorkbook newBook = new XSSFWorkbook();
             FileOutputStream outputStream = new FileOutputStream(destFilePath)) {
            XSSFSheet sheet = wb.getSheet("easy");
            XSSFSheet destSheet = newBook.createSheet("test");
            int dataListRow = 0;
            // 先列后行，这样可以方便生成
            for (int columnIndex = 0; columnIndex < 4; columnIndex++) {
                // 遍历每个单元格
                HashMap<Integer, String> hashMap = new HashMap<>();
                HashMap<String, Map> map = new HashMap<>();
                ObjectNode variableNode = JsonNodeFactory.instance.objectNode();
                List<List<String>> lists = new ArrayList<>();
                for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    XSSFRow row = sheet.getRow(rowIndex);
                    if (row != null) {
                        Cell cell = row.getCell(columnIndex);

                        if (cell != null) {
                            String value = getCellValue(cell);

                            // 变量，需要替换
                            if (value.startsWith("{")) {

                                String variable = value.substring(1, value.length() - 1);

                                if (variable.contains("[")) {
                                    dataListRow = rowIndex;


                                    ArrayList<String> strings = new ArrayList<>();
                                    JsonNode className = var.get("className");
                                    for (JsonNode jsonNode : className) {
                                        String s = jsonNode.asText();
                                        strings.add(s);
                                    }
                                    lists.add(strings);
                                } else {
                                    String[] split = variable.split("\\.");
                                    String firstVariable = split[0];
                                    String others = variable.replaceFirst(firstVariable + ".", "");

                                    // 表示数组
                                    hashMap.put(rowIndex, firstVariable);
                                    JsonNode arr = var.get(firstVariable);
                                    if (arr == null) {
                                        continue;
                                    }
                                    if (arr.isArray()) {
                                        ArrayList<String> strings = new ArrayList<>();
                                        for (JsonNode jsonNode : arr) {
                                            String text = null;
                                            if (jsonNode.isObject()) {
                                                ObjectNode tempNode = JsonNodeFactory.instance.objectNode();
                                                tempNode.put(firstVariable, jsonNode);
                                                variableNode.put(String.valueOf(fastIndex++), tempNode);
//                                            HashMap temp = new HashMap<String, Object>();
//                                            temp.put(firstVariable, jsonNode);
//                                            map.put(String.valueOf(fastIndex++), temp);
                                                text = jsonNode.findPath(others).asText();
                                            } else if (jsonNode.isTextual()) {
                                                text = jsonNode.asText();
                                            }
                                            strings.add(text);
                                        }
                                        lists.add(strings);
                                        /*for (JsonNode jsonNode : arr) {

                                            // hashMap.put()
//                                            map.put(rowIndex,)
                                            // 每次自动copy一列
                                            ExcelUtil.copyColumn(sheet, destSheet, columnIndex, destColumnIndex);
                                            destSheet.getRow(rowIndex).getCell(destColumnIndex++).setCellValue(text);
                                        }*/
//                                        wb.write(outputStream);
//                                        return;
                                    }
                                    log.info("{} ---", firstVariable);

                                    log.info(variable);
                                }
//                                System.out.println(variable);
                            } else {
                                lists.add(List.of(value));
                            }


                            // 输出单元格内容
                            log.info("[{}, {}]: {}", rowIndex, columnIndex, value);
//                            System.out.println("[" + rowIndex + "," + columnIndex + "]: " + value);
                        }
                    }
                }

                int index = 0;

                iter(lists, new LinkedList<>(), index, sheet, destSheet, columnIndex, dataListRow, variableNode, data);

                newBook.write(outputStream);
            }
            outputStream.close();
            wb.close();
        } catch (
                Exception e) {
            throw new Exception(step + "生成Excel发生错误！", e);
        }

    }

    private void iter(List<List<String>> lists, LinkedList<String> list, int index, Sheet sourceSheet, Sheet targetSheet, int sourceColumnIndex, int dataListRow, JsonNode node, JsonNode data) {
        if (index == dataListRow) {
            ExcelUtils.copyColumn(sourceSheet, targetSheet, sourceColumnIndex, ind++, list, dataListRow, node, data);
//            targetColumnIndex++;

            System.out.println(list.toString() + "-" + ind);
            return;
        }
        List<String> strings = lists.get(index);
        for (String string : strings) {
            list.add(string);
            iter(lists, list, index + 1, sourceSheet, targetSheet, sourceColumnIndex, dataListRow, node, data);
            list.pollLast();
        }
    }

    private static String getCellValue(Cell cell) {
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
}