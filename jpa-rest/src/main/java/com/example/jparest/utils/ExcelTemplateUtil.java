package com.example.jparest.utils;

import com.example.jparest.utils.ExcelUtils;
import com.example.jparest.utils.FileDownloadUtil;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.ClassPathResource;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Component;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.util.*;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

@Slf4j
@Component
@Data
public class ExcelTemplateUtil {
    private int writeColumnIndex = 0; // write dest excel column index
    private int mergeColumnIndex = 0; // merge dest excel column index
    private List<List<JsonNode>> variableNodes = new ArrayList<>(); // 每列所用的变数集合，index行，value：该行的用到的变量集合
    private int dataListRow;

    @Autowired
    private FileDownloadUtil fileDownloadUtil;


    public ResponseEntity<ByteArrayResource> generateExcel(List<?> dataList, String projectName, Map<String, ?> appendVars) {
        return generateMultiSheetExcel(List.of(dataList), projectName, List.of(appendVars));
    }

    public <T, F> ResponseEntity<ByteArrayResource> generateMultiSheetExcel(List<List<T>> dataList, String projectName, List<Map<String, F>> appendVarsList) {
        String step = "";
        step = "[1.0].Read template";

        ObjectMapper mapper = JsonUtils.mapper;


        try (InputStream templateStream = new ClassPathResource(ExcelUtils.TEMPLATE + File.separator + projectName + "File.xlsx").getInputStream();
             InputStream inputStream = new ClassPathResource(ExcelUtils.JSON + File.separator + projectName + "Vars.json").getInputStream();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
             XSSFWorkbook wb = new XSSFWorkbook(templateStream);
             XSSFWorkbook newBook = new XSSFWorkbook();) {

            JsonNode variables = mapper.readTree(inputStream);


            XSSFSheet sheet = wb.getSheetAt(0);

            for (int i = 0; i < dataList.size(); i++) {
                List<T> data = dataList.get(i);
                Map<String, F> appendVars = appendVarsList.get(i);
                JsonNode dataNode = mapper.valueToTree(data);
                ObjectNode nodes = mapper.valueToTree(appendVars);
                ((ObjectNode) variables).setAll(nodes);

                step = "[2.0].Read the template and start writing Excel";

                XSSFSheet newSheet = newBook.createSheet();

                fillTemplate(sheet, newSheet, variables, dataNode);
                ZipSecureFile.setMinInflateRatio(-1.0d);
            }

            newBook.write(outputStream);
            return fileDownloadUtil.downloadFile(outputStream, projectName);

        } catch (Exception e) {
            throw new MyExceptionUtil(step + "生成Excel发生错误！", e);
        }
    }

    // copy column from one sheet to another
    public void copyColumn(Sheet sourceSheet, Sheet targetSheet, int sourceColumnIndex, int targetColumnIndex, List<String> list, int dataListRow, LinkedList<JsonNode> node, JsonNode data, HashMap<String, Integer> varLineMap) throws Exception {
        String sourceValue = sourceSheet.getRow(dataListRow).getCell(sourceColumnIndex).getStringCellValue();
        if (!sourceValue.isEmpty()) {
            sourceValue = sourceValue.substring(1, sourceValue.length() - 1);
        }
        String newPath = Pattern.compile("\\[(.*?)\\]").matcher(sourceValue).replaceAll(m -> {
            String group = m.group(1);
            // 在前面解决好路径的问题
            group = group.replaceAll("\\.", "/");
            String[] split = group.split("/");
            String variable = split[0];
            String collect = "/" + Arrays.stream(split).skip(1).collect(Collectors.joining("/"));
//            collect =
            if (!varLineMap.containsKey(variable)) {
                throw new NullPointerException(String.format("%s is not exist in column variables", variable));
            }
            Integer integer = varLineMap.get(variable);
            JsonNode jsonNode = node.get(integer);
            String text;
            if (jsonNode.isObject()) {
                JsonNode at = jsonNode.at(collect);
                if (at.isMissingNode()) {
                    throw new NullPointerException(String.format("%s is not exist in variables %s", collect, jsonNode));
                } else {
                    text = at.asText();
                }
            } else {
                text = jsonNode.asText();
            }
            return "/" + text;
        });

        // to:　data/sub/cost/point/10
        newPath = newPath.replaceAll("\\.", "/");
        // 单字符可选可不选
        if (!newPath.isBlank()) {
            if (newPath.startsWith("//")) {
                newPath = newPath.substring(1);
            } else {
                // to: /sub/cost/point/10
                newPath = newPath.substring(newPath.indexOf("/"));
            }
        }
        List<String> strings = new ArrayList<>();
        if (data.isArray()) {
            for (JsonNode item : data) {
                JsonNode at = item.at(newPath);
                if (at.isMissingNode() || at.isNull()) {
                    strings.add("");
                } else {
                    strings.add(at.asText());
                }
            }
        } else {
            throw new Exception("data is not array");
        }

        for (int i = 0; i < dataListRow + strings.size(); i++) {
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


    public void fillTemplate(XSSFSheet sheet, XSSFSheet destSheet, JsonNode variables, JsonNode data) throws Exception {
        dataListRow = 0;
        writeColumnIndex = 0;
        mergeColumnIndex = 0;
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
                Cell cell = row.getCell(columnIndex);
                if (cell == null) {
                    break;
                }
                String value = ExcelUtils.getCellValue(cell);

                readVariables(variables, value, rowIndex, varIndexMap, lists, placeholderNodes);
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

    public void readVariables(JsonNode variables, String value, int rowIndex, HashMap<String, Integer> varIndexMap, List<List<String>> lists, List<JsonNode> placeholderNodes) throws Exception {
        // 变量，需要替换
        if (value.startsWith("{")) {

            String variable = value.substring(1, value.length() - 1);

            if (variable.contains("[")) {
                dataListRow = rowIndex;
                // fix: 单行表头的时候报错
                lists.add(List.of(value));
            } else {
                ArrayList<String> strings = new ArrayList<>();
                ArrayList<JsonNode> tempNodes = new ArrayList<>();
                String[] split = variable.split("\\.");
                findValue(variables, rowIndex, varIndexMap, strings, tempNodes, split, 0);
                variableNodes.add(tempNodes);
                lists.add(strings);
            }
        } else {
            variableNodes.add(placeholderNodes);
            lists.add(List.of(value));
        }
    }

    void findValue(JsonNode variable, int rowIndex, HashMap<String, Integer> varIndexMap, ArrayList<String> strings, ArrayList<JsonNode> tempNodes, String[] split, int index) throws Exception {
        findValue(variable, rowIndex, varIndexMap, strings, tempNodes, split, index, "");
    }

    private void findValue(JsonNode variable, int rowIndex, HashMap<String, Integer> varIndexMap, ArrayList<String> strings, ArrayList<JsonNode> tempNodes, String[] split, int index, String path) throws Exception {
        // 表示数组
        if (variable == null) {
            throw new Exception(path + " is not exist in variables");
        }
        boolean isLast = index == split.length - 1;
        boolean isLastPlus1 = index == split.length;
        if (isLastPlus1) {
            if (variable.isValueNode()) {
                strings.add(variable.asText());
            } else if (variable.isArray()) {
                // 字串数组，就都保存起来
                boolean areAllValue = true;
                ArrayList<String> temp = new ArrayList<>();
                ArrayList<JsonNode> tempNode = new ArrayList<>();

                for (JsonNode jsonNode : variable) {
                    if (jsonNode.isValueNode()) {
                        tempNode.add(jsonNode);
                        temp.add(jsonNode.asText());
                    } else {
                        areAllValue = false;
                    }
                }
                // 先清空，再加上temp nodes
                if (areAllValue) {
                    tempNodes.clear();
                    tempNodes.addAll(tempNode);
                    strings.addAll(temp);
                }
            }
            return;
        }
        String var = split[index];
//        else
        if (index == 0) {
            varIndexMap.put(var, rowIndex);
        }

        if (variable.isArray()) {
            int i = 0;
            for (JsonNode jsonNode : variable) {
                if (isLast) {
                    tempNodes.add(jsonNode);
                }
                findValue(jsonNode.get(var), rowIndex, varIndexMap, strings, tempNodes, split, index + 1, path + "/" + (i++) + "/" + var);

            }
        } else {
            if (isLast) {
                tempNodes.add(variable);
            }
            findValue(variable.get(var), rowIndex, varIndexMap, strings, tempNodes, split, index + 1, path + "/" + var);
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
                                      data, HashMap<String, Integer> varLineMap) throws Exception {
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
