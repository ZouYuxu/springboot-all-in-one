package com.example.jparest.service;

import cn.hutool.core.io.FileUtil;
import com.example.jparest.utils.ExcelUtils;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
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
    private List<List<JsonNode>> variableNodes;

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
//        File source = new ClassPathResource("templates/easy.xlsx").getFile();
//        File dest = new ClassPathResource("templates/dest.xlsx").getFile();
        JsonNode var = mapper.readTree(inputStream);
        JsonNode data = mapper.readTree(dataInputStream);
        String step = "[1.0].Read template";
        String sourceFilePath = "D:\\workspace\\java\\springboot-all-in-one\\jpa-rest\\src\\main\\resources\\templates\\easy.xlsx";
        String destFilePath = "D:\\workspace\\java\\springboot-all-in-one\\jpa-rest\\src\\main\\resources\\templates\\dest.xlsx";
        try (XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(sourceFilePath));
             XSSFWorkbook newBook = new XSSFWorkbook();
             FileOutputStream outputStream = new FileOutputStream(destFilePath)) {
            XSSFSheet sheet = wb.getSheet("easy");
            XSSFSheet destSheet = newBook.createSheet("what");
            int dataListRow = 0;
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            // 先列后行，这样可以方便生成
            for (int columnIndex = 0; columnIndex < sheet.getRow(0).getLastCellNum(); columnIndex++) {
                int finalColumnIndex = columnIndex;
                Optional<CellRangeAddress> rangeAddress = mergedRegions.stream().filter(v -> v.containsColumn(finalColumnIndex)).findFirst();
                // 遍历每个单元格
                HashMap<Integer, String> hashMap = new HashMap<>();
                HashMap<String, Integer> map = new HashMap<>();
                JsonNode variableNode = JsonNodeFactory.instance.objectNode();
                List<JsonNode> placeholderNodes = List.of(variableNode);
                List<List<String>> lists = new ArrayList<>();
                variableNodes = new ArrayList<>();
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


//                                    ArrayList<String> strings = new ArrayList<>();
//                                    JsonNode className = var.get("className");
//                                    for (JsonNode jsonNode : className) {
//                                        String s = jsonNode.asText();
//                                        strings.add(s);
//                                    }
//                                    lists.add(strings);
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
                                        ArrayList<JsonNode> tempNodes = new ArrayList<>();
                                        for (JsonNode jsonNode : arr) {
                                            String text;
//                                            ObjectNode tempNode = JsonNodeFactory.instance.objectNode();
                                            if (jsonNode.isObject()) {
//                                            HashMap temp = new HashMap<String, Object>();
//                                            temp.put(firstVariable, jsonNode);
//                                            map.put(String.valueOf(fastIndex++), temp);
                                                text = jsonNode.findPath(others).asText();
                                            } else {
                                                text = jsonNode.asText();
                                            }
                                            map.put(firstVariable, rowIndex);
                                            tempNodes.add(jsonNode);
                                            strings.add(text);
                                        }
                                        variableNodes.add(tempNodes);
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
                                    } else {
                                        variableNodes.add(placeholderNodes);
                                        lists.add(List.of(value));
                                    }
                                    log.info("{} ---", firstVariable);

                                    log.info(variable);
                                }
//                                System.out.println(variable);
                            } else {
                                variableNodes.add(placeholderNodes);
                                lists.add(List.of(value));
                            }


                            // 输出单元格内容
                            log.info("[{}, {}]: {}", rowIndex, columnIndex, value);
//                            System.out.println("[" + rowIndex + "," + columnIndex + "]: " + value);
                        }
                    }
                }

                int index = 0;
                int initial = this.ind;

                iter(lists, new LinkedList<>(), index, sheet, destSheet, columnIndex, dataListRow, new LinkedList<>(), data, map);


                createMerge(0, lists, dataListRow, columnIndex, destSheet, initial, rangeAddress);
                System.out.println();
                if (columnIndex == 10) {

                    newBook.write(outputStream);
                    return;
                }
//                destSheet.addMergedRegion(new CellRangeAddress(0, 2, 0, 0));
            }
            newBook.write(outputStream);
            outputStream.close();
            wb.close();
        } catch (
                Exception e) {
            throw new Exception(step + "生成Excel发生错误！", e);
        }

    }

    private int createMerge(int rowIndex, List<List<String>> lists, int dataListRow, int columnIndex, XSSFSheet destSheet, int initial, Optional<CellRangeAddress> rangeAddress) {
        // 获取下一个列表的长度
        List<String> rowVars = lists.get(rowIndex);
        int rowSize = rowVars.size();

        if (rowIndex == dataListRow) return rowSize;

        int nextSize = createMerge(rowIndex + 1, lists, dataListRow, columnIndex, destSheet, initial, rangeAddress);
        int newSize = rowSize * nextSize;


        if (rangeAddress.isPresent()) {
            CellRangeAddress cellAddresses = rangeAddress.get();
            int firstRangeColumn = cellAddresses.getFirstColumn();
            int lastRangeColumn = cellAddresses.getLastColumn();
            int firstRangeRow = cellAddresses.getFirstRow();
            int lastRangeRow = cellAddresses.getLastRow();
            int delta = lastRangeColumn - firstRangeColumn;

            // 单个元素不需要分组
//            if (!cellAddresses.containsRow(rowIndex) && nextSize == 1 && delta == 1) {
//                return nextSize;
//            }


            for (int sizeIndex = 0; sizeIndex < rowSize; sizeIndex++) {
                int firstCol = initial + sizeIndex * nextSize; // 0
                int lastCol = firstCol + nextSize - 1; // 2
                //


                // 不包含，就不处理
                // 第一行
                if (firstRangeRow == rowIndex) {
//                            destSheet.addMergedRegion(new CellRangeAddress(firstRangeRow, lastRangeRow, ))

                    // 包含但是不是第一行，就不处理
                } else if (cellAddresses.containsRow(rowIndex)) {
                    return newSize;
                } else {
                    firstRangeRow = rowIndex;
                    lastRangeRow = rowIndex;

                }
                // 列
                // 第一列
                // 如果是跨列的话，只取第一个进行merge，防止重复
                boolean isColspan = delta > 1;
                if (isColspan) {
                    // 不适第一个
                    if (sizeIndex != 0) {
                        // 结束循环
                        break;
                    }
                    firstRangeColumn = firstCol;
                    lastRangeColumn = nextSize * delta + lastCol;
                } else {
                    // 不跨列
                    firstRangeColumn = firstCol;
                    lastRangeColumn = lastCol;

                }
                log.info("{}-{}: {}", rowIndex, sizeIndex, nextSize);
                if (firstRangeRow == lastRangeRow && firstRangeColumn == lastRangeColumn) {
                    continue;
                }
                destSheet.addMergedRegion(new CellRangeAddress(firstRangeRow, lastRangeRow, firstRangeColumn, lastRangeColumn));
            }

        } else {
            if (nextSize > 1) {
                for (int sizeIndex = 0; sizeIndex < rowSize; sizeIndex++) {

                    int firstCol = initial + sizeIndex * nextSize; // 0
                    int lastCol = firstCol + nextSize - 1; // 2

                    destSheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, firstCol, lastCol));
                    log.info("{}-{}: {}", rowIndex, sizeIndex, nextSize);
                }

//                List<String> next2List = lists.get(rowIndex + 2);
//                int next2Size = next2List.size();
//                while (next2Size > 1) {
//
//                }
//
//                List<String> row = lists.get(rowIndex);
//                for (int sizeIndex = 0; sizeIndex < row.size(); sizeIndex++) {
//                    int nextSize = lists.get(rowIndex + 1).size();
//                    if (nextSize > 1) {
//
//                    }
//                    int firstCol = initial + sizeIndex * nextSize; // 0
//                    int lastCol = firstCol + nextSize - 1; // 2
//
//                    destSheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, firstCol, lastCol));
//                    log.info("{}-{}: {}", rowIndex, sizeIndex, nextSize);
//                }

            }
        }

        return newSize;

    }


    private void iter(List<List<String>> lists, LinkedList<String> list, int index, Sheet sourceSheet, Sheet targetSheet, int sourceColumnIndex, int dataListRow, LinkedList<JsonNode> variables, JsonNode data, HashMap<String, Integer> varLineMap) {
        if (index == dataListRow) {
            ExcelUtils.copyColumn(sourceSheet, targetSheet, sourceColumnIndex, ind++, list, dataListRow, variables, data, varLineMap);
//            targetColumnIndex++;

            System.out.println(list.toString() + "-" + ind);
            return;
        }
        List<String> strings = lists.get(index);
        List<JsonNode> jsonNodes = variableNodes.get(index);
        for (int i = 0; i < strings.size(); i++) {
            String string = strings.get(i);
            JsonNode jsonNode = jsonNodes.get(i);
            list.add(string);
            variables.add(jsonNode);
            iter(lists, list, index + 1, sourceSheet, targetSheet, sourceColumnIndex, dataListRow, variables, data, varLineMap);
            list.pollLast();
            variables.pollLast();
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