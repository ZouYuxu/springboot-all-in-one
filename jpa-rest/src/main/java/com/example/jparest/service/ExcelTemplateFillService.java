package com.example.jparest.service;

import cn.hutool.core.io.resource.ResourceUtil;
import com.example.jparest.utils.ExcelTemplateUtil;
import com.example.jparest.utils.FileDownloadUtil;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.util.ObjectUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;

@Service
public class ExcelTemplateFillService {
    public ResponseEntity<ByteArrayResource> fillTemplate(MultipartFile template, String vars, String data) throws IOException {
        if (ObjectUtils.isEmpty(template)) {
            throw new FileNotFoundException("必须上传文件！");
        }
        ObjectMapper mapper = new ObjectMapper();
        JsonNode varsNode = mapper.readTree(vars);
        JsonNode dataNode = mapper.readTree(data);
        try (XSSFWorkbook wb = new XSSFWorkbook(template.getInputStream());
             XSSFWorkbook newBook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            XSSFSheet sheet = wb.getSheet("easy");
            XSSFSheet newSheet = newBook.createSheet();
            ExcelTemplateUtil excelTemplateUtil = new ExcelTemplateUtil();
            excelTemplateUtil.fillTemplate(sheet, newSheet, varsNode, dataNode);
            newBook.write(outputStream);
            return new FileDownloadUtil().downloadFile(outputStream, "new");
        } catch (Exception e) {
            // 处理异常情况
            e.printStackTrace();
            throw new RuntimeException("文件格式错误！");
        }
    }
    public void fillTemplate(MultipartFile template, String vars, String data, HttpServletResponse response) throws IOException {
        if (ObjectUtils.isEmpty(template)) {
            throw new FileNotFoundException("必须上传文件！");
        }
        ObjectMapper mapper = new ObjectMapper();
        JsonNode varsNode = mapper.readTree(vars);
        JsonNode dataNode = mapper.readTree(data);
        try (XSSFWorkbook wb = new XSSFWorkbook(template.getInputStream());
             XSSFWorkbook newBook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            XSSFSheet sheet = wb.getSheet("easy");
            XSSFSheet newSheet = newBook.createSheet();
            ExcelTemplateUtil excelTemplateUtil = new ExcelTemplateUtil();
            excelTemplateUtil.fillTemplate(sheet, newSheet, varsNode, dataNode);
            //准备将Excel的输出流通过response输出到页面下载
            //八进制输出流
            response.setContentType("application/octet-stream");
//            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            //这后面可以设置导出Excel的名称，此例中名为student.xls
            response.setHeader("Content-disposition", String.format("attachment;filename=%s.xlsx","demo"+new Date().getTime()));

            //刷新缓冲
            response.flushBuffer();

            //workbook将Excel写入到response的输出流中，供页面下载
            newBook.write(response.getOutputStream());
//            newBook.write(outputStream);
        } catch (Exception e) {
            // 处理异常情况
            e.printStackTrace();
            throw new RuntimeException("文件格式错误！");
        }
    }

    public void fillTemplate(HttpServletResponse response) throws IOException {

        ObjectMapper mapper = new ObjectMapper();
        JsonNode varsNode = mapper.readTree(ResourceUtil.getStream("json/var.json"));
        JsonNode dataNode = mapper.readTree(ResourceUtil.getStream("json/data.json"));
        try (XSSFWorkbook wb = new XSSFWorkbook(ResourceUtil.getStream("template/EvaluationAnalyseFile.xlsx"));
             XSSFWorkbook newBook = new XSSFWorkbook();
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFSheet newSheet = newBook.createSheet();
            ExcelTemplateUtil excelTemplateUtil = new ExcelTemplateUtil();
            excelTemplateUtil.fillTemplate(sheet, newSheet, varsNode, dataNode);
            //准备将Excel的输出流通过response输出到页面下载
            //八进制输出流
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            //这后面可以设置导出Excel的名称，此例中名为student.xls
            response.setHeader("Content-disposition", "attachment;filename=employee.xlsx");

            //刷新缓冲
            response.flushBuffer();

            //workbook将Excel写入到response的输出流中，供页面下载
            newBook.write(response.getOutputStream());
//            newBook.write(outputStream);
        } catch (Exception e) {
            // 处理异常情况
            e.printStackTrace();
            throw new RuntimeException("文件格式错误！");
        }
    }
}
