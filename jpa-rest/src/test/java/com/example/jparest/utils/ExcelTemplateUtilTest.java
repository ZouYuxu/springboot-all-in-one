package com.example.jparest.utils;

import cn.hutool.core.io.resource.ResourceUtil;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;
import org.slf4j.Logger;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

import static com.example.jparest.service.EvaluationAnalyseDownloadService.DEST_FILE_PATH;
import static org.junit.jupiter.api.Assertions.*;

class ExcelTemplateUtilTest {
    public static final String EASY = "easy";
    public static final String EASY_1 = "easy_1";
    public static final String EASY_2 = "easy_2";
    public static final String EASY_3 = "easy_3";
    List<List<JsonNode>> variableNodes;
    @Mock
    Logger log;
    @InjectMocks
    ExcelTemplateUtil excelTemplateUtil;

    ObjectMapper mapper = new ObjectMapper();
    private JsonNode variables;
    private JsonNode array;
    HashMap<String, Integer> varIndexMap = new HashMap<>();
    List<JsonNode> placeholderNodes = List.of(JsonNodeFactory.instance.objectNode());
    List<List<String>> lists = new ArrayList<>();
    private InputStream easy;
    private InputStream evaluationContentAgg;
    private InputStream error;
    private JsonNode data;
    private JsonNode evaluationContentAggData;
    private XSSFWorkbook errorWb;
    private XSSFWorkbook wb;

    @BeforeEach
    void setUp() throws IOException {
        MockitoAnnotations.openMocks(this);

        variables = mapper.readTree(ResourceUtil.getStreamSafe("json/var.json"));
        array = mapper.readTree(ResourceUtil.getStreamSafe("json/test.json"));
        data = mapper.readTree(ResourceUtil.getStreamSafe("json/data.json"));
        evaluationContentAggData = mapper.readTree(ResourceUtil.getStreamSafe("json/EvaluationContentAgg_data.json"));
        variableNodes = new ArrayList<>();
        easy = ResourceUtil.getStream("template/easy.xlsx");
        evaluationContentAgg = ResourceUtil.getStream("templates/EvaluationContentAggFile.xlsx");
        error = ResourceUtil.getStream("template/error.xlsx");

        errorWb = new XSSFWorkbook(error);
        wb = new XSSFWorkbook(easy);
        System.out.println();
    }

    //    @Test
    void testCopyColumn() throws Exception {
        excelTemplateUtil.copyColumn(null, null, 0, 0, List.of("list"), 0, new LinkedList<>(List.of(null)), null, new HashMap<>(Map.of("varLineMap", Integer.valueOf(0))));
    }

    @Test
    void readVariables_const() throws Exception {
        excelTemplateUtil.readVariables(variables, "整合", 0, varIndexMap, lists, placeholderNodes);
        List<String> strings = lists.get(0);
        assertEquals(List.of("整合"), strings);
    }

    @Test
    void readVariables() throws Exception {
        excelTemplateUtil.readVariables(variables, "{class.key}", 0, varIndexMap, lists, placeholderNodes);
        List<String> strings = lists.get(0);
        assertEquals(List.of("cost", "delivery", "quality", "esg", "design"), strings);
        assertEquals(0, varIndexMap.get("class"));
    }

    @Test
    void readVariables_notExist_throwException() throws Exception {
        assertThrows(Exception.class, () -> excelTemplateUtil.readVariables(variables, "{class.keys}", 0, varIndexMap, lists, placeholderNodes), "/class/0/keys is not exist in variables");
        assertTrue(lists.isEmpty());
        assertTrue(excelTemplateUtil.getVariableNodes().isEmpty());
        assertEquals(0, varIndexMap.get("class"));
    }

    @Test
    void readSingleValue() throws Exception {
        excelTemplateUtil.readVariables(variables, "{keyName}", 0, varIndexMap, lists, placeholderNodes);
        List<String> strings = lists.get(0);
        List<List<JsonNode>> variableNodes = excelTemplateUtil.getVariableNodes();
        assertEquals(List.of("work"), strings);
        assertEquals(0, varIndexMap.get("keyName"));
    }

    @Test
    void readMultipleValue() throws Exception {
        excelTemplateUtil.readVariables(variables, "{level1.level2}", 0, varIndexMap, lists, placeholderNodes);
        List<String> strings = lists.get(0);
        List<List<JsonNode>> variableNodes = excelTemplateUtil.getVariableNodes();
        assertEquals(List.of("yes"), strings);
        assertEquals(0, varIndexMap.get("level1"));
    }

    @Test
    void findValue() throws Exception {
        String[] split = new String[]{"level1", "level2"};
        ArrayList<String> strings = new ArrayList<>();
        ArrayList<JsonNode> tempNodes = new ArrayList<>();
        excelTemplateUtil.findValue(variables, 0, varIndexMap, strings, tempNodes, split, 0);
//        List<String> strings = lists.get(0);
        List<List<JsonNode>> variableNodes = excelTemplateUtil.getVariableNodes();
        assertEquals(List.of("yes"), strings);
        assertEquals(0, varIndexMap.get("level1"));
    }

    @Test
    void findStringArrayValue() throws Exception {
        String[] split = new String[]{"mon"};
        ArrayList<String> strings = new ArrayList<>();
        ArrayList<JsonNode> tempNodes = new ArrayList<>();
        excelTemplateUtil.findValue(variables, 0, varIndexMap, strings, tempNodes, split, 0);
//        List<String> strings = lists.get(0);
        List<List<JsonNode>> variableNodes = excelTemplateUtil.getVariableNodes();
        assertEquals(List.of("10", "11", "12"), strings);
        assertEquals(0, varIndexMap.get("mon"));
    }

    @Test
    void findArrayValue() throws Exception {
        String[] split = new String[]{"data", "key"};
        ArrayList<String> strings = new ArrayList<>();
        ArrayList<JsonNode> tempNodes = new ArrayList<>();
        excelTemplateUtil.findValue(array, 0, varIndexMap, strings, tempNodes, split, 0);
//        List<String> strings = lists.get(0);
        List<List<JsonNode>> variableNodes = excelTemplateUtil.getVariableNodes();
        assertEquals(List.of("quality", "esg", "quality", "esg", "quality", "esg"), strings);
        assertEquals(0, varIndexMap.get("data"));
    }

    @Test
    void findArrayKey() throws Exception {
        String[] split = new String[]{"key"};
        ArrayList<String> strings = new ArrayList<>();
        ArrayList<JsonNode> tempNodes = new ArrayList<>();
        excelTemplateUtil.findValue(array, 0, varIndexMap, strings, tempNodes, split, 0);
//        List<String> strings = lists.get(0);
        List<List<JsonNode>> variableNodes = excelTemplateUtil.getVariableNodes();
        assertEquals(List.of("delivery", "quality", "esg"), strings);
        assertEquals(0, varIndexMap.get("key"));
    }

    @Test
    void findArrayNestKey() throws Exception {
        String[] split = new String[]{"object", "key"};
        ArrayList<String> strings = new ArrayList<>();
        ArrayList<JsonNode> tempNodes = new ArrayList<>();
        excelTemplateUtil.findValue(array, 0, varIndexMap, strings, tempNodes, split, 0);
//        List<String> strings = lists.get(0);
        List<List<JsonNode>> variableNodes = excelTemplateUtil.getVariableNodes();
        assertEquals(List.of("delivery", "delivery", "delivery"), strings);
        assertEquals(0, varIndexMap.get("object"));
    }

    @Test
    void testFillTemplate() throws Exception {
        try (XSSFWorkbook newBook = new XSSFWorkbook();
             FileOutputStream outputStream = new FileOutputStream(DEST_FILE_PATH)) {
            XSSFSheet sheet = wb.getSheet(EASY);
            XSSFSheet destSheet = newBook.createSheet();

            ExcelTemplateUtil excelTemplateUtil = new ExcelTemplateUtil();
            excelTemplateUtil.fillTemplate(sheet, destSheet, variables, data);
            newBook.write(outputStream);
            outputStream.close();
            System.out.println();
        } catch (
                Exception e) {
            throw new Exception("生成Excel发生错误！", e);
        }
    }

    @Test
    void testFillEvaluationContentAgg() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(evaluationContentAgg);
             XSSFWorkbook newBook = new XSSFWorkbook();
             FileOutputStream outputStream = new FileOutputStream(DEST_FILE_PATH + ".xlsx")) {
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFSheet destSheet = newBook.createSheet();

            ExcelTemplateUtil excelTemplateUtil = new ExcelTemplateUtil();
            excelTemplateUtil.fillTemplate(sheet, destSheet, variables, evaluationContentAggData);
            newBook.write(outputStream);
            outputStream.close();
            System.out.println();
        } catch (
                Exception e) {
            throw new Exception("生成Excel发生错误！", e);
        }
    }

    @Test
    void testFillErrorTemplate()  {
//        XSSFWorkbook newBook = new XSSFWorkbook();
//        FileOutputStream outputStream = new FileOutputStream(DEST_FILE_PATH);
        XSSFSheet sheet = errorWb.getSheetAt(0);
        XSSFSheet sheet1 = errorWb.getSheetAt(1);

        XSSFSheet destSheet = errorWb.createSheet();

        ExcelTemplateUtil excelTemplateUtil = new ExcelTemplateUtil();
//            excelTemplateUtil.fillTemplate(sheet, destSheet, variables, evaluationContentAggData);
        assertThrows(Exception.class, () -> excelTemplateUtil.fillTemplate(sheet, destSheet, variables, evaluationContentAggData), "/keys is not exist in variables {\"name\":\"brand\",\"key\":\"brand\"}");
        assertThrows(Exception.class, () -> excelTemplateUtil.fillTemplate(sheet1, destSheet, variables, evaluationContentAggData),"dfsa");

    }
}

//Generated with love by TestMe :) Please report issues and submit feature requests at: https://weirddev.com/forum#!/testme