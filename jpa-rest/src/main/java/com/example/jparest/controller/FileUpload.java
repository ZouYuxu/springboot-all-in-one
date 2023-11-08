package com.example.jparest.controller;

import com.example.jparest.utils.FileUtil;
import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController
@RequiredArgsConstructor
public class FileUpload {

    @PostMapping("file")
    public String saveFile(@RequestParam("file") MultipartFile file) throws IOException {
        return FileUtil.saveFile(file);
    }
}
