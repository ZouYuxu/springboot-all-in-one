package com.example.jparest.service;

import cn.hutool.core.io.FileUtil;
import com.example.jparest.utils.ExcelTemplateUtil;
import com.example.jparest.utils.ExcelUtils;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;

import java.io.*;

import static cn.hutool.poi.excel.cell.CellUtil.getCellValue;


@Slf4j
public class EvaluationAnalyseDownloadService {

    public static final String SOURCE_FILE_PATH = "E:\\workspace\\java\\springboot-all-in-one\\jpa-rest\\src\\main\\resources\\templates\\easy.xlsx";
    public static final String DEST_FILE_PATH = "E:\\workspace\\java\\springboot-all-in-one\\jpa-rest\\src\\main\\resources\\templates\\dest.xlsx";
    public static final String EASY = "easy";
    public static final String WHAT = "what";

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

//        File source = new ClassPathResource("templates/easy.xlsx").getFile();
//        File dest = new ClassPathResource("templates/dest.xlsx").getFile();

    public void template() throws Exception {
        ObjectMapper mapper = new ObjectMapper();
        InputStream inputStream = new ClassPathResource("json/var.json").getInputStream();
        InputStream dataInputStream = new ClassPathResource("json/data.json").getInputStream();
        JsonNode variables = mapper.readTree(inputStream);
        JsonNode data = mapper.readTree(dataInputStream);
        String step = "[1.0].Read template";
        try (XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(SOURCE_FILE_PATH));
             XSSFWorkbook newBook = new XSSFWorkbook();
             FileOutputStream outputStream = new FileOutputStream(DEST_FILE_PATH)) {
            XSSFSheet sheet = wb.getSheet(EASY);
            XSSFSheet destSheet = newBook.createSheet(WHAT);

            ExcelTemplateUtil excelTemplateUtil = new ExcelTemplateUtil();
            excelTemplateUtil.fillTemplate(sheet, destSheet, variables, data);

            newBook.write(outputStream);
        } catch (
                Exception e) {
            throw new Exception(step + "生成Excel发生错误！", e);
        }

    }
}