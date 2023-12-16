package com.example.jparest;

import cn.hutool.core.convert.Convert;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.resource.ResourceUtil;
import cn.hutool.core.lang.TypeReference;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.example.jparest.handler.CustomCellWriteHandler;
import com.example.jparest.model.Book;
import com.example.jparest.model.request.BookCreationRequest;
import com.example.jparest.utils.ExcelUtil;
import com.example.jparest.utils.TestFileUtil;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.ClassPathResource;

import java.io.*;
import java.util.List;

class DownloadTest {

    @Test
    public void cloneSheet() throws IOException {
        String result = TestFileUtil.getPath() + "template/easyResult.xlsx";
        String template = TestFileUtil.getPath() + "template/easy.xlsx";
        ObjectMapper mapper = new ObjectMapper();
        JsonNode jsonNode = mapper.readTree(FileUtil.readUtf8String("js.json"));
        EasyExcel.write(result)
                .registerWriteHandler(new CustomCellWriteHandler())
                .withTemplate(template)
                .sheet("easy")
                .doFill(jsonNode);
    }
}
