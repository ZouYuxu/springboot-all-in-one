package org.example.domain;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

import java.util.List;

@Component
@ConfigurationProperties(prefix = "user")
@Data
public class User {
    private String name;
    private Integer age;
    // 可以是数组，也可以是List
    private List<String> aliases;
}
