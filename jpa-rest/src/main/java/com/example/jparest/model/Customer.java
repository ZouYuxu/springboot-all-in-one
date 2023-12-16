package com.example.jparest.model;

import lombok.Data;

import java.util.Date;

@Data
public class Customer {
    private String name;
    private int authorId;
    private int id;
    private String  isbn;
    private Double bonus;
    private String  now;
}