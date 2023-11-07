package com.example.jparest.service;

import com.example.jparest.model.Book;
import com.example.jparest.model.request.BookCreationRequest;
import com.example.jparest.repository.AuthorRepository;
import com.example.jparest.repository.BookRepository;
import com.example.jparest.repository.LendRepository;
import com.example.jparest.repository.MemberRepository;
import jakarta.persistence.EntityNotFoundException;
import org.springframework.beans.BeanUtils;
import org.springframework.stereotype.Service;

import lombok.RequiredArgsConstructor;

import java.util.List;
import java.util.Optional;

@Service
@RequiredArgsConstructor
public class LibraryService {

    private final AuthorRepository authorRepository;
    private final MemberRepository memberRepository;
    private final LendRepository lendRepository;
    private final BookRepository bookRepository;

    public List<Book> readBooks() {
        return bookRepository.findAll();
    }

    public Book createBook(BookCreationRequest book) {
//        Optional<Book> byId = bookRepository.findByIsbn(book.getIsbn());
//        if (byId.isPresent()) {
//            return
//        }
        Book bookToC = new Book();
        BeanUtils.copyProperties(book, bookToC);
        return bookRepository.save(bookToC);

    }


    public Book updateBook(BookCreationRequest request) {
        Optional<Book> byId = bookRepository.findById(request.getId());
        if (!byId.isPresent()) {
            throw new EntityNotFoundException("Member not find !");
        }
        Book bookToC = new Book();
        BeanUtils.copyProperties(request, bookToC);
        return bookRepository.save(bookToC);

    }
}