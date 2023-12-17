package com.example.jparest;

import ch.qos.logback.classic.turbo.MarkerFilter;
import cn.hutool.core.convert.Convert;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.resource.ResourceUtil;
import com.example.jparest.model.Book;
import com.example.jparest.model.Customer;
import com.example.jparest.model.request.BookCreationRequest;
import com.example.jparest.utils.ExcelUtil;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.extern.slf4j.Slf4j;
import org.apache.logging.log4j.Marker;
import org.apache.logging.log4j.MarkerManager;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.platform.commons.logging.LoggerFactory;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.springframework.core.io.ClassPathResource;

import java.io.*;
import java.util.List;

@Slf4j
class JpaRestApplicationTests {

    ObjectMapper mapper = new ObjectMapper();
    String filePath = "template" + File.separator + "a.xlsx";
    String jsonPath = "customers.json";
    String destFilePath = "D:\\workspace\\java\\springboot-all-in-one\\jpa-rest\\src\\test\\resources\\template\\dest.xlsx";
    String sourFilePath = "D:\\workspace\\java\\springboot-all-in-one\\jpa-rest\\src\\test\\resources\\template\\a.xlsx";

    @Test
    public void jxls() throws IOException {
        List<Customer> employees = generateSampleEmployeeData();
        try (InputStream is = ResourceUtil.getStreamSafe(filePath);
             OutputStream os = new FileOutputStream(destFilePath)) {
            Context context = new Context();
            context.putVar("customers", generateSampleEmployeeData());
            JxlsHelper.getInstance().processTemplate(is, os, context);
        }
    }

    private List<Customer> generateSampleEmployeeData() throws JsonProcessingException {
        String s = FileUtil.readUtf8String(jsonPath);
        List<Customer> customers = mapper.readValue(s, new TypeReference<List<Customer>>() {
        });
        return customers;
    }

    @Test
    public void cloneSheet() throws IOException {
        String filePath2 = "template" + File.separator + "easy1.xlsx";

        byte[] bytes = FileUtil.readBytes(filePath);

//        InputStream resourceAsStream = this.getClass().getResourceAsStream(filePath);
//        byte[] bytes = resourceAsStream.readAllBytes();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes));
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream();) {
            XSSFSheet mock = wb.getSheet("mock");
            XSSFSheet sheet = null;
            if (mock == null) {

                sheet = wb.cloneSheet(0, "mock");
            }
//            XSSFSheet sheet = wb.getSheet("easy");
//            wb.sethi
            ExcelUtil.copyColumn(sheet, 0, 0 + 1);
//            String stringCellValue = sheet.getRow(0).getCell(0).getStringCellValue();
//            System.out.println(stringCellValue);


            wb.write(outputStream);
            FileUtil.writeBytes(outputStream.toByteArray(), destFilePath);
            outputStream.close();
            wb.close();
//            wb.write(fileOutputStream);
//            FileUtil.writeToStream(filePath2, outputStream);

            outputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    public void downloadEvaluationAnalyse() throws Exception {
        ObjectMapper mapper = new ObjectMapper();
        InputStream inputStream = new ClassPathResource("var.json").getInputStream();
        JsonNode var = mapper.readTree(inputStream);

        String step = "[1.0].Read template";
        byte[] bytes = FileUtil.readBytes("template" + File.separator + "easy.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes));
             FileOutputStream outputStream = new FileOutputStream(destFilePath)) {
            XSSFSheet sheet = wb.getSheet("easy");
            // 先列后行，这样可以方便生成
            for (int columnIndex = 0; columnIndex < sheet.getRow(0).getLastCellNum(); columnIndex++) {
                // 遍历每个单元格
                for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    XSSFRow row = sheet.getRow(rowIndex);

                    if (row != null) {
                        Cell cell = row.getCell(columnIndex);

                        if (cell != null) {
                            String value = getCellValue(cell);

                            // 变量，需要替换
                            if (value.startsWith("{")) {

                                String variable = value.substring(1, value.length() - 1);
                                // 表示数组
                                if (variable.startsWith("...#")) {
                                    String newStr = variable.replaceFirst("...#", "");
                                    JsonNode arr = var.get(newStr);
                                    if (arr == null) {
                                        continue;
                                    }
                                    if (arr.isArray()) {
                                        int i = 0;
                                        for (JsonNode jsonNode : arr) {
                                            String text = jsonNode.asText();
                                            if (i == 0) {
                                                i += 1;
                                                cell.setCellValue(text);

                                            } else {
//                                                sheet.shiftColumns(columnIndex, sheet.getRow(0).getLastCellNum() -1, 1);

//                                                ExcelUtil.copyColumn(sheet, columnIndex, columnIndex + 1, text);
                                            }
//                                            System.out.println(text + " 0_0");
                                            log.info("++ {}", text);
                                        }
                                    }
                                    log.info("{} ---", newStr);
                                }
                                log.info(variable);
//                                System.out.println(variable);
                            }


                            // 输出单元格内容
                            log.info("[{}, {}]: {}", rowIndex, columnIndex, value);
//                            System.out.println("[" + rowIndex + "," + columnIndex + "]: " + value);
                        }
                    }
                }

            }
                wb.write(outputStream);
//            outputStream.close();
            wb.close();
        } catch (Exception e) {
            throw new Exception(step + "生成Excel发生错误！", e);
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

    @Test
    void t() throws JsonProcessingException {
        ObjectMapper mapper = new ObjectMapper();
        String s = ResourceUtil.readUtf8Str("category.json");
        // 替换变量为字串(可变数组类型)
        String newStr = s.replaceFirst("#plants", "F131.F132.F721");
        JsonNode jsonNode = mapper.readTree(newStr);
        extracted(jsonNode, "-");
    }

    private void extracted(JsonNode jsonNode, String size) {
        System.out.println(size);
        jsonNode.fields().forEachRemaining((entry) -> {
            String key = entry.getKey();
            JsonNode node = entry.getValue();

            // 不需要作为表头
            if (key.startsWith("__")) {
                System.out.println("hello");
            }


            // 以.间隔表示重复处理key
            String[] split = key.split("\\.");
            if (split.length > 1) {
                for (String s1 : split) {

                }
            }

            // 嵌套
            if (node.isObject()) {
                extracted(node, size + "-" + key);
            }


            System.out.println(key);
        });
    }

    @Test
    void contextLoads() throws IOException {
        // 读取 resources 下面的json文件
        File test = new File("src/test/resources/customers.json");
        File main = new File("src/main/resources/customers.json");
        InputStream reader = ResourceUtil.getStreamSafe("customers.json");
//		InputStream in = getClass().getResourceAsStream("/customers.json");
        File in = new ClassPathResource("customers.json").getFile();

//		BufferedReader bufferedReader = Files.newBufferedReader(Paths.get("src/main/resources/customers.json"));

        // "jackson"
        // json 转 vo
        ObjectMapper mapper = new ObjectMapper();
        BookCreationRequest bookCreationRequest = mapper.readValue(reader, BookCreationRequest.class);
        // vo 转json
        String s = mapper.writeValueAsString(bookCreationRequest);

        // "hutool":　
        // bean copy
        Book convert = Convert.convert(Book.class, bookCreationRequest);
        // 转换为自己想要的类型的类型
//        List<Book> converts = Convert.convert(new TypeReference<>() {
//        }, List.of(bookCreationRequest));

        System.out.println(convert);
//        System.out.println(converts);
        System.out.println(bookCreationRequest);
        System.out.println(s);
    }

}
