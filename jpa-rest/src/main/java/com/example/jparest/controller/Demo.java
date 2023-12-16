package com.example.jparest.controller;

import com.example.jparest.model.Book;
import com.example.jparest.model.request.BookCreationRequest;
import com.example.jparest.service.LibraryService;
import lombok.RequiredArgsConstructor;
import lombok.SneakyThrows;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;

@RestController
@RequiredArgsConstructor
public class Demo {
    private final LibraryService libraryService;

    @GetMapping("books")
    public ResponseEntity readBooks() {
        return ResponseEntity.ok(libraryService.readBooks());
    }

    @PostMapping("books")
    public ResponseEntity<Book> createBook(@RequestBody BookCreationRequest request) {

        Book book = libraryService.createBook(request);
        return ResponseEntity.ok(book);
    }

    @PutMapping("books")
    @SneakyThrows
    public ResponseEntity<Book> updateBook(@RequestBody BookCreationRequest request) {
        Book book = libraryService.updateBook(request);
        return ResponseEntity.ok(book);
    }
}
