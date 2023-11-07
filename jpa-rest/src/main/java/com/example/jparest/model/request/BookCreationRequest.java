package com.example.jparest.model.request;

import jakarta.persistence.Id;
import lombok.Data;

@Data
public class BookCreationRequest {
    private Long id;
    private String name;
    private String isbn;
    private Long authorId;
}