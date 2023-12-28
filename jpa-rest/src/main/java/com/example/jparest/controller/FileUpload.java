package com.example.jparest.controller;

import cn.hutool.core.io.resource.ResourceUtil;
import com.example.jparest.service.ExcelTemplateFillService;
import com.example.jparest.utils.FileUtil;
import com.fasterxml.jackson.databind.ObjectMapper;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController
@RequiredArgsConstructor
public class FileUpload {

    @Autowired
    ExcelTemplateFillService excelTemplateFillService;

    @PostMapping("file")
    public String saveFile(@RequestParam("file") MultipartFile file) throws IOException {
        return FileUtil.saveFile(file);
    }

    @GetMapping("excel")
    public void save() {
        // 写法2
        String fileName = "simpleWrite" + ".xlsx";
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        // 如果这里想使用03 则 传入excelType参数即可

        ObjectMapper mapper = new ObjectMapper();
        String s = ResourceUtil.readUtf8Str("category.json");
        System.out.println(s);
//        EasyExcel.write(fileName, DemoData.class).sheet("模板").doWrite(data());
    }

    @PostMapping(value = "/fillTemplate")
    public ResponseEntity<ByteArrayResource> fillTemplate(MultipartFile file, String vars, String data) throws IOException {
        return excelTemplateFillService.fillTemplate(file,vars,data);
    }

    @PostMapping(value = "/fillExcel")
    public void fillTemplate(MultipartFile file, String vars, String data, HttpServletResponse response) throws IOException {
        excelTemplateFillService.fillTemplate(file,vars,data,response);
    }

    @GetMapping(value = "/fill")
    public void fillTemplate(HttpServletResponse response) throws IOException {
        excelTemplateFillService.fillTemplate(response);
    }
}
