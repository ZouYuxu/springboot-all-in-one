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

import static org.junit.jupiter.api.Assertions.assertEquals;

class ExcelTemplateUtilTest {
    List<List<JsonNode>> variableNodes;
    @Mock
    Logger log;
    @InjectMocks
    ExcelTemplateUtil excelTemplateUtil;

    ObjectMapper mapper = new ObjectMapper();
    private JsonNode variables;
    HashMap<String, Integer> varIndexMap = new HashMap<>();
    List<JsonNode> placeholderNodes = List.of(JsonNodeFactory.instance.objectNode());
    List<List<String>> lists = new ArrayList<>();

    @BeforeEach
    void setUp() throws IOException {
        MockitoAnnotations.openMocks(this);

        InputStream stream = ResourceUtil.getStreamSafe("json/var.json");
        variables = mapper.readTree(stream);
        variableNodes = new ArrayList<>();
    }

    @Test
    void testCopyColumn() {
        excelTemplateUtil.copyColumn(null, null, 0, 0, List.of("list"), 0, new LinkedList<>(List.of(null)), null, new HashMap<>(Map.of("varLineMap", Integer.valueOf(0))));
    }

    @Test
    void readVariables() {
        excelTemplateUtil.readVariables(variables, "{class.key}", 0, varIndexMap, lists, placeholderNodes);
        List<String> strings = lists.get(0);
        assertEquals(List.of("cost", "delivery", "quality", "esg", "design"), strings);
        assertEquals(0, varIndexMap.get("class"));
    }

    @Test
    void readVariabless() {
        excelTemplateUtil.readVariables(variables, "{class.keys}", 0, varIndexMap, lists, placeholderNodes);
        List<String> strings = lists.get(0);
        assertEquals(List.of("cost", "delivery", "quality", "esg", "design"), strings);
        assertEquals(0, varIndexMap.get("class"));
    }

    @Test
    void readSingleValue() {
        excelTemplateUtil.readVariables(variables, "{keyName}", 0, varIndexMap, lists, placeholderNodes);
        List<String> strings = lists.get(0);
        List<List<JsonNode>> variableNodes = excelTemplateUtil.getVariableNodes();
        assertEquals(List.of("work"), strings);
        assertEquals(0, varIndexMap.get("keyName"));
    }

    @Test
    void readMultipleValue() {
        excelTemplateUtil.readVariables(variables, "{level1.level2}", 0, varIndexMap, lists, placeholderNodes);
        List<String> strings = lists.get(0);
        List<List<JsonNode>> variableNodes = excelTemplateUtil.getVariableNodes();
        assertEquals(List.of("work"), strings);
        assertEquals(0, varIndexMap.get("keyName"));
    }

    @Test
    void testFillTemplate() {
        excelTemplateUtil.fillTemplate(null, null, null, null);
    }
}

//Generated with love by TestMe :) Please report issues and submit feature requests at: https://weirddev.com/forum#!/testme