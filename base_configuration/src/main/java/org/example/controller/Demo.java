package org.example.controller;

import lombok.extern.slf4j.Slf4j;
import org.example.domain.User;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.env.Environment;
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

    @Autowired
    private Environment env;

    @Autowired
    private User user;
    @GetMapping
    public String hello() {
        log.info(name);
        log.info(port);
        log.info(arrs);
        log.info(baseDir);
        log.info(tempDir);
        log.info("------");
        log.info(env.getProperty("baseDir"));
        log.info(String.valueOf(user));
        for (String alias : user.getAliases()) {
            log.info(alias);
        }
        return "Hello World!";
    }
}
