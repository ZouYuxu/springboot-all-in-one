package com.example.jparest;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.resource.ResourceUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.enums.WriteDirectionEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.example.jparest.handler.CustomCellWriteHandler;
import com.example.jparest.handler.CustomConverterHandler;
import com.example.jparest.service.EvaluationAnalyseDownloadService;
import com.example.jparest.utils.TestFileUtil;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.util.*;

class DownloadTest {
    ObjectMapper mapper = new ObjectMapper();

    // target/test-classes/template
    // 路径
    String result = TestFileUtil.getPath() + "template/easyResult.xlsx";
    String template = TestFileUtil.getPath() + "template/easy.xlsx";
    EvaluationAnalyseDownloadService evaluationAnalyseDownloadService = new EvaluationAnalyseDownloadService();
    @Test
    void tt() throws Exception {
        evaluationAnalyseDownloadService.template();
    }

    @Test
    void t() {
        ExcelReader reader = ExcelUtil.getReader(template, "easy");
        List<Object> easy = reader.readColumn(0, 0);
        Object o = reader.readCellValue(0, 1);
//        reader.rea
        System.out.println();
//        reader.getWriter().write
//        ExcelUtil.
    }

    @Test
    void excelWrite() {
        ExcelWriter excelWriter = null;
        try {
            // 封装表头蓝色区域数据
            HashMap map = mapper.readValue(FileUtil.readUtf8String("qu.json"),HashMap.class);

            // 存放表头数据&填充数据需要的花括号数据
            List<Map> headList = new ArrayList<>();
            // 存放结果集数据
            List<Map> dataList = new ArrayList<>();
            for (int i = 0; i < 10; i++) {
                // 封装表头数据&花括号数据
                Map headMap = new HashMap(16);
                headMap.put("list1", "list" + i);
                headMap.put("list2", "{data2.heihei" + i + "}");
                headList.add(headMap);
                // 封装结果集数据
                Map dataMap = new HashMap(16);
                dataMap.put("heihei0", "heihei0name" + i);
                dataMap.put("heihei1", "heihei1name" + i);
                dataMap.put("heihei2", "heihei2name" + i);
                dataMap.put("heihei3", "heihei3name" + i);
                dataMap.put("heihei4", "heihei4name" + i);
                dataMap.put("heihei5", "heihei5name" + i);
                dataMap.put("heihei6", "heihei6name" + i);
                dataMap.put("heihei7", "heihei7name" + i);
                dataMap.put("heihei8", "heihei8name" + i);
                dataMap.put("heihei9", "heihei9name" + i);
                dataList.add(dataMap);
            }
            // 这里我是存放到resources下的temp文件夹内 自己修改文件路径

            //
            InputStream inputStream = ResourceUtil.getStream(template);
//            File test = File.createTempFile("test", ".xlsx");
//            FileUtils.copyInputStreamToFile(inputStream, test);
//            response.setContentType("application/vnd.ms-excel");
//            response.addHeader("Access-Control-Expose-Headers", "Content-Disposition");
//            response.setCharacterEncoding("utf-8");
//            response.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode("test", "UTF-8") + ".xlsx");
//            ServletOutputStream outputStream = response.getOutputStream();

            // 处理合并单元格 我这里是将第一列蓝色区域长度跟下面保持一致 合并成一个单元格
//            FileUtil.createTempFile();
            FileUtil.touch(result);
            excelWriter = EasyExcel.write(result).withTemplate(template)
                    // 因为这个合并单元格，导致我资料格式错误
//                    .registerWriteHandler(new AbstractMergeStrategy() {
//                @Override
//                protected void merge(Sheet sheet, Cell cell, Head head, Integer relativeRowIndex) {
//                    sheet.addMergedRegionUnsafe(new CellRangeAddress(0, 0, 0, headList.size() - 1));
//                }
//            })
                    .registerConverter(new CustomConverterHandler())
                    .registerWriteHandler(new CustomCellWriteHandler()).build();
            // 注意sheet名字要改成自己的模板sheet
            WriteSheet writeSheet = EasyExcel.writerSheet("easy").build();
            // 指定横向填充
            FillConfig fillConfig = FillConfig.builder().direction(WriteDirectionEnum.HORIZONTAL).build();
            // 横向填充表头以及花括号数据
//            excelWriter.fill(new FillWrapper("base", (List) map.get("base")), fillConfig, writeSheet);
//            // 纵向填充结果集数据
//            excelWriter.fill(new FillWrapper("data2", dataList), writeSheet);
            // 填充蓝色区域数据
            excelWriter.fill(map, writeSheet);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }

    }

    @Test
    public void cloneSheet() throws IOException {
        String result = TestFileUtil.getPath() + "template/easyResult.xlsx";
        String template = TestFileUtil.getPath() + "template/easy.xlsx";
        ObjectMapper mapper = new ObjectMapper();
        HashMap jsonNode = mapper.readValue(FileUtil.readUtf8String("jss.json"),HashMap.class);


        HashMap<String, String> map = new HashMap<>();
        map.put("name", "张三");
        map.put("number", "5.2");
        EasyExcel.write(result)
                .registerWriteHandler(new CustomCellWriteHandler())
                .withTemplate(template)
                .sheet("easy")
                .doFill(map);
    }
}
