package com.example.jparest.utils;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.*;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

@Slf4j
public class ExcelTemplateUtil {
    private static int writeColumnIndex = 0; // write dest excel column index
    private static int mergeColumnIndex = 0; // merge dest excel column index
    private static List<List<JsonNode>> variableNodes; // 每列所用的变数集合，index行，value：该行的用到的变量集合

    // copy column from one sheet to another
    public void copyColumn(Sheet sourceSheet, Sheet targetSheet, int sourceColumnIndex, int targetColumnIndex, List<String> list, int dataListRow, LinkedList<JsonNode> node, JsonNode data, HashMap<String, Integer> varLineMap) {
        int i = 0;
        List<String> ss = List.of("我真的", "kusi");

        for (; i < dataListRow; i++) {
            boolean flag = i >= dataListRow;
            int index = i - dataListRow;
            Row sourceRow = sourceSheet.getRow(flag ? dataListRow : i);
//            sourceSheet.addMergedRegion(new CellRangeAddress())
            Row targetRow = targetSheet.getRow(i);
            String newValue = flag ? ss.get(index) : list.get(i);
            // fix bug：brand column missing
            if (targetRow == null) {
                targetRow = targetSheet.createRow(i);
            }
            if (sourceRow != null) {
                Cell sourceCell = sourceRow.getCell(sourceColumnIndex);
                Cell newCell = targetRow.createCell(targetColumnIndex);
                ExcelUtils.copyCell(sourceCell, newCell, newValue);
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
            String collect = "/" + Arrays.stream(split).skip(1).collect(Collectors.joining("/"));
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
                JsonNode at = item.at(newPath);
                strings.add(at.asText());
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
                Cell newCell = targetRow.createCell(targetColumnIndex);
                ExcelUtils.copyCell(sourceCell, newCell, newValue);
            }
        }
    }


    public void fillTemplate(XSSFSheet sheet, XSSFSheet destSheet, JsonNode variables, JsonNode data) {
        int dataListRow = 0;
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        // 先列后行，这样可以方便生成
        for (int columnIndex = 0; columnIndex < sheet.getRow(0).getLastCellNum(); columnIndex++) {
            int finalColumnIndex = columnIndex;
            Optional<CellRangeAddress> rangeAddress = mergedRegions.stream().filter(v -> v.containsColumn(finalColumnIndex)).findFirst();
            // 遍历每个单元格
            HashMap<String, Integer> varIndexMap = new HashMap<>();
            JsonNode variableNode = JsonNodeFactory.instance.objectNode();
            List<JsonNode> placeholderNodes = List.of(variableNode);
            List<List<String>> lists = new ArrayList<>();
            variableNodes = new ArrayList<>();
            for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                XSSFRow row = sheet.getRow(rowIndex);
                if (row != null) {
                    Cell cell = row.getCell(columnIndex);

                    if (cell != null) {
                        String value = ExcelUtils.getCellValue(cell);

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
                                JsonNode arr = variables.get(firstVariable);
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
                                        varIndexMap.put(firstVariable, rowIndex);
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

            createColumn(lists, new LinkedList<>(), 0, sheet, destSheet, columnIndex, dataListRow, new LinkedList<>(), data, varIndexMap);

            createMerge(0, lists, dataListRow, columnIndex, destSheet, rangeAddress);
        }

        // 自动列宽
        for (int i = 0; i < writeColumnIndex; i++) {
            destSheet.autoSizeColumn(i, true);
            destSheet.setColumnWidth(i, destSheet.getColumnWidth(i) * 11 / 10);
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
        int temp = mergeColumnIndex;
        // 不是初始列，不需要
        if (!isStartColumn) {
            return newSize;
        }
        for (int sizeIndex = 0; sizeIndex < rowVarSize; sizeIndex++) {
            int nextSize = createMerge(rowIndex + 1, lists, dataListRow, columnIndex, destSheet, rangeAddress);
            newSize = rowVarSize * nextSize;

            if (nextSize > 1) {
                int firstCol = mergeColumnIndex; // 0
                // 增加间隔
                int lastCol = firstCol + nextSize * (1 + columnGaps) - 1; // 2
                mergeColumnIndex = lastCol + 1;
                // 跳过为空的，不合并
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
                    destSheet.addMergedRegion(new CellRangeAddress(rowIndex, lastRowIndex, mergeColumnIndex, mergeColumnIndex));
                }
                // 单列自动增加1
                mergeColumnIndex += 1;
            }
        }
        // 如果是第一行的话，是真的会影响到initial
        if (rowIndex != 0) {
            mergeColumnIndex = temp;
        }
        return newSize;
    }

    private void createColumn(List<List<String>> lists, LinkedList<String> list, int index, Sheet sourceSheet, Sheet
            targetSheet, int sourceColumnIndex, int dataListRow, LinkedList<JsonNode> variables, JsonNode
                                      data, HashMap<String, Integer> varLineMap) {
        if (index == dataListRow) {
            copyColumn(sourceSheet, targetSheet, sourceColumnIndex, writeColumnIndex++, list, dataListRow, variables, data, varLineMap);
            return;
        }
        List<String> strings = lists.get(index);
        List<JsonNode> jsonNodes = variableNodes.get(index);
        for (int i = 0; i < strings.size(); i++) {
            String string = strings.get(i);
            JsonNode jsonNode = jsonNodes.get(i);
            list.add(string);
            variables.add(jsonNode);
            createColumn(lists, list, index + 1, sourceSheet, targetSheet, sourceColumnIndex, dataListRow, variables, data, varLineMap);
            list.pollLast();
            variables.pollLast();
        }
    }


}