package com.example.jparest.utils;

import cn.hutool.core.io.resource.ResourceUtil;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;
import org.slf4j.Logger;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

import static org.junit.jupiter.api.Assertions.*;

class ExcelTemplateUtilTest {
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

    @BeforeEach
    void setUp() throws IOException {
        MockitoAnnotations.openMocks(this);

        InputStream stream = ResourceUtil.getStreamSafe("json/var.json");
        variables = mapper.readTree(stream);
        array = mapper.readTree(ResourceUtil.getStreamSafe("json/test.json"));
        variableNodes = new ArrayList<>();
    }

//    @Test
    void testCopyColumn() {
        excelTemplateUtil.copyColumn(null, null, 0, 0, List.of("list"), 0, new LinkedList<>(List.of(null)), null, new HashMap<>(Map.of("varLineMap", Integer.valueOf(0))));
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
        assertEquals(List.of("delivery","delivery","delivery"), strings);
        assertEquals(0, varIndexMap.get("object"));
    }

//    @Test
    void testFillTemplate() throws Exception {
        excelTemplateUtil.fillTemplate(null, null, null, null);
    }
}

//Generated with love by TestMe :) Please report issues and submit feature requests at: https://weirddev.com/forum#!/testme