package com.example.jparest.model.service;


import com.example.jparest.model.entity.LikedSongs;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.Pageable;

import java.util.Collection;
import java.util.List;


/**
 * 业务层
 *
 * @author makejava
 * @since 2023-11-12 10:10:59
 */
public interface LikedSongsService {

    void save(LikedSongs likedSongs);

    void deleteById(Long id);

    LikedSongs findById(Long id);

    List<LikedSongs> findById(Collection<Long> ids);

    Page<LikedSongs> list(Pageable page);

}

