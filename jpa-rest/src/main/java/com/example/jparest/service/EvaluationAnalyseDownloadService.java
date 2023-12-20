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
    private int initial;

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
        String sourceFilePath = "E:\\workspace\\java\\springboot-all-in-one\\jpa-rest\\src\\main\\resources\\templates\\easy.xlsx";
        String destFilePath = "E:\\workspace\\java\\springboot-all-in-one\\jpa-rest\\src\\main\\resources\\templates\\dest.xlsx";
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
                                            if (jsonNode.isObject()) {
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
                                    } else {
                                        variableNodes.add(placeholderNodes);
                                        lists.add(List.of(value));
                                    }
                                    log.info("{} --- {}", firstVariable, variable);
                                }
                            } else {
                                variableNodes.add(placeholderNodes);
                                lists.add(List.of(value));
                            }
                            // 输出单元格内容
                            log.info("[{}, {}]: {}", rowIndex, columnIndex, value);
                        }
                    }
                }

                int index = 0;
                int initial = this.ind;

                iter(lists, new LinkedList<>(), index, sheet, destSheet, columnIndex, dataListRow, new LinkedList<>(), data, map);


                createMerge(0, lists, dataListRow, columnIndex, destSheet, rangeAddress);
            }

            // 自动列宽
            for (int i = 0; i < ind; i++) {
                destSheet.autoSizeColumn(i, true);
                destSheet.setColumnWidth(i, destSheet.getColumnWidth(i) * 11 / 10);
            }
            newBook.write(outputStream);
            outputStream.close();
            wb.close();
        } catch (
                Exception e) {
            throw new Exception(step + "生成Excel发生错误！", e);
        }

    }


    private int createMerge(int rowIndex, List<List<String>> lists, int dataListRow, int columnIndex, XSSFSheet destSheet, Optional<CellRangeAddress> rangeAddress) {
        List<String> rowVars = lists.get(rowIndex);
        int rowVarSize = rowVars.size();

        if (rowIndex == dataListRow) return rowVarSize;
        boolean isStartColumn = true;
        int lastRowIndex = rowIndex;
        int columnGaps = 0;
        if (rangeAddress.isPresent()) {
            CellRangeAddress cellAddresses = rangeAddress.get();
            // 包含row的话
            if (cellAddresses.containsRow(rowIndex)) {
                lastRowIndex = cellAddresses.getLastRow();
            }

            // 包含col的话
            if (cellAddresses.containsColumn(columnIndex)) {
                columnGaps = cellAddresses.getLastColumn() - cellAddresses.getFirstColumn();
            }

            isStartColumn = cellAddresses.getFirstColumn() == columnIndex;
        }
        int newSize = 1;
        int temp = initial;
        // 不是初始列，不需要
        if (!isStartColumn) {
            return newSize;
        }
        for (int sizeIndex = 0; sizeIndex < rowVarSize; sizeIndex++) {
            int nextSize = createMerge(rowIndex + 1, lists, dataListRow, columnIndex, destSheet, rangeAddress);
            newSize = rowVarSize * nextSize;

            if (nextSize > 1) {
                int firstCol = initial; // 0
                // 增加间隔
                int lastCol = firstCol + nextSize * (1 + columnGaps) - 1; // 2
                initial = lastCol + 1;
                // todo: 跳过为空的不合并
                if (rowIndex != 0 && rowVarSize == 1 && rowVars.get(0).isEmpty()) {
                    break;
                }
                destSheet.addMergedRegion(new CellRangeAddress(rowIndex, lastRowIndex, firstCol, lastCol));
                log.info("{}-{}: {}", rowIndex, sizeIndex, nextSize);
            } else {
                if (rowIndex != 0 && rowVarSize == 1 && rowVars.get(0).isEmpty()) {
                    break;
                }
                // 防止单个单元格合并报错
                if (rowIndex != lastRowIndex) {
                    destSheet.addMergedRegion(new CellRangeAddress(rowIndex, lastRowIndex, initial, initial));
                }
                // 单列自动增加1
                initial += 1;
            }
        }
        // 如果是第一行的话，是真的会影响到initial
        if (rowIndex != 0) {
            initial = temp;
        }
        return newSize;
    }

    private void iter(List<List<String>> lists, LinkedList<String> list, int index, Sheet sourceSheet, Sheet
            targetSheet, int sourceColumnIndex, int dataListRow, LinkedList<JsonNode> variables, JsonNode
                              data, HashMap<String, Integer> varLineMap) {
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