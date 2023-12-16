//package com.example.jparest.model.service.impl;
//
//
//import com.example.jparest.model.entity.LikedSongs;
//import com.example.jparest.model.repository.LikedSongsRepository;
//import com.example.jparest.model.service.LikedSongsService;
//import org.springframework.data.domain.Page;
//import org.springframework.data.domain.Pageable;
//import org.springframework.stereotype.Service;
//
//import javax.annotation.Resource;
//import java.util.Collection;
//import java.util.List;
//import java.util.stream.Collectors;
//import java.util.stream.StreamSupport;
//
//
///**
// * 业务层
// *
// * @author makejava
// * @since 2023-11-12 10:11:02
// */
//@Service
//public class LikedSongsServiceImpl implements LikedSongsService {
//
//	@Resource
//    private LikedSongsRepository likedSongsRepository;
//
//    @Override
//    public void save(LikedSongs likedSongs) {
//        likedSongsRepository.save(likedSongs);
//    }
//
//    @Override
//    public void deleteById(Long id) {
//        likedSongsRepository.delete(id);
//    }
//
//    @Override
//    public LikedSongs findById(Long id) {
//        return likedSongsRepository.findOne(id);
//    }
//
//    @Override
//    public List<LikedSongs> findById(Collection<Long> ids) {
//        Iterable<LikedSongs> iterable = likedSongsRepository.findAll(ids);
//        return StreamSupport.stream(iterable.spliterator(), false)
//                .collect(Collectors.toList());
//    }
//
//    @Override
//    public Page<LikedSongs> list(Pageable page) {
//        return likedSongsRepository.findAll(page);
//    }
//
//}
//
