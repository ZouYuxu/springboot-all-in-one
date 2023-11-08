package com.example.jparest.utils;

import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;

public class FileUtil {
    public static String saveFile(MultipartFile file) throws IOException {
        String originalFilename = file.getOriginalFilename();
        String filePath = "D:/file/";

        File dest = new File(filePath + originalFilename);
//        File parentFile = dest.getParentFile();
//        if (!parentFile.exists()) {
//            parentFile.mkdirs();
//        }
        file.transferTo(dest);
        return dest.getAbsolutePath();
    }
}
