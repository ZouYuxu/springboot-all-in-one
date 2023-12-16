package com.example.jparest.model.request;

import com.fasterxml.jackson.annotation.JsonFormat;
import jakarta.persistence.Id;
import lombok.Data;

import java.util.Date;

@Data
public class BookCreationRequest {
    private Long id;
    private String name;
    private String isbn;
    private Long authorId;

    @JsonFormat(pattern = "yyyy-MM-dd HH:mm:ss", timezone = "GMT+8")
    private Date now;
}