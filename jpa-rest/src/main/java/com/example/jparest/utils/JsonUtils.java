package com.example.jparest.utils;


import cn.hutool.core.io.resource.ResourceUtil;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.IOException;

public class JsonUtils {
    public static ObjectMapper mapper = new ObjectMapper();

//    public JsonUtils() {
//        mapper.setPropertyNamingStrategy(PropertyNamingStrategies.SNAKE_CASE);
//    }

    public static String read(String filename) {
        return ResourceUtil.readUtf8Str(filename);
    }

    public static JsonNode readTree(String filePath) throws IOException {
        return mapper.readTree(ResourceUtil.getStreamSafe(filePath));
    }

    public static <T> T fromJsonFile(String filename, Class<T> clazz) throws JsonProcessingException {
        return mapper.readValue(read(filename), clazz);
    }

    public static <T> T fromJsonFile(String filename, TypeReference<T> typeReference) throws JsonProcessingException {
        return mapper.readValue(read(filename), typeReference);
    }

    public static <T> T fromJson(String json, Class<T> clazz) throws JsonProcessingException {
        return mapper.readValue(json, clazz);
    }

    public static <T> T fromJson(String json, TypeReference<T> typeReference) throws JsonProcessingException {
        return mapper.readValue(json, typeReference);
    }

    public static <T> String toJson(T obj) throws JsonProcessingException {
        return mapper.writeValueAsString(obj);
    }

}
