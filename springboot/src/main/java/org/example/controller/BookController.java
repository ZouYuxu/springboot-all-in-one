package org.example.controller;

import lombok.extern.slf4j.Slf4j;
import org.example.domain.Book;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("books")
@Slf4j
public class BookController {
    @PostMapping
    public Book createBook(@RequestBody Book book) {
        log.info("Creating book: {}", book);
        return book;
    }

    @DeleteMapping("/{id}")
    public void deleteBook(@PathVariable("id") Long id) {
        log.info("Deleting book with id: {}", id);
    }

    @PutMapping
    public Book updateBook(@RequestBody Book book) {
        log.info("Updating book: {}", book);
        return book;
    }
    @GetMapping("/{id}")
    public Book getBook(@PathVariable("id") Long id) {
        log.info("Getting book with id: {}", id);
        return new Book();
    }

    @GetMapping
    public List<Book> getAllBooks() {
        log.info("Getting all books");
        return null;
    }
}
