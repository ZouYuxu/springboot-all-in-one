package org.example.controller;

import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@Slf4j
public class Demo {
    @Value("name")
    private String name;

    @Value("${server.port}")
    private String port;

    @Value("${arrs[1]}")
    private String arrs;

    @Value("${baseDir}")
    private String baseDir;

    @Value("${tempDir}")
    private String tempDir;

    @GetMapping
    public String hello() {
        log.info(name);
        log.info(port);
        log.info(arrs);
        log.info(baseDir);
        log.info(tempDir);
        return "Hello World!";
    }
}
