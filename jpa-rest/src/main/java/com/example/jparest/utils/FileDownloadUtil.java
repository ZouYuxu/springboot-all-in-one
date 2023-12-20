package com.example.jparest.utils;

import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Component;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.net.URLEncoder;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
@Component
public class FileDownloadUtil {

    public ResponseEntity<ByteArrayResource> downloadFile(ByteArrayOutputStream outputStream,
                                                          String fileName)
            throws IOException {
        String date = this.getNowTime();
        String strFileName = this.getFileName(date, fileName);  // 创建文件名

        // 设置下载文件名编码，防止中文文件名乱码
        String encodedFilename = URLEncoder.encode(strFileName, "UTF-8").replace("\\+", "%20");

        // 将ByteArrayOutputStream转换为byte数组
        byte[] byteArray = outputStream.toByteArray();

        // 创建一个资源对象，将byte数组作为参数传递给构造函数
        ByteArrayResource resource = new ByteArrayResource(byteArray);

        // 创建一个ResponseEntity对象，将资源对象作为参数传递给构造函数，并设置HTTP状态码和头部信息
        HttpHeaders headers = new HttpHeaders();
        headers.add(HttpHeaders.CONTENT_DISPOSITION,
                String.format("attachment; fileName=%s", encodedFilename));
        headers.set(HttpHeaders.CONTENT_TYPE, "application/octet-stream");
        ResponseEntity<ByteArrayResource> responseEntity = new ResponseEntity<>(
                resource, headers, HttpStatus.OK);

        return responseEntity;
    }

    private String getFileName(String date, String filename) {
        return String.format("%s%s.xlsx", filename, date);
    }

    private String getNowTime() {
        LocalDate currentDate = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy_MM_dd");
        return currentDate.format(formatter);
    }
}
