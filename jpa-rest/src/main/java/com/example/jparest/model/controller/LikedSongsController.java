//package com.example.jparest.model.controller;
//
//
//import com.example.jparest.model.entity.LikedSongs;
//import com.example.jparest.model.service.LikedSongsService;
//import lombok.AllArgsConstructor;
//import org.springframework.beans.factory.annotation.Autowired;
//import org.springframework.data.domain.Page;
//import org.springframework.data.domain.Pageable;
//import org.springframework.web.bind.annotation.*;
//
//
///**
// * 控制层
// *
// * @author makejava
// * @since 2023-11-12 10:10:49
// */
//@RestController
//@RequestMapping("/likedSongs")
//@AllArgsConstructor
//public class LikedSongsController {
//
//	private LikedSongsService likedSongsService;
//
//	/**
//	 * 获取列表(分页)
//	 */
//	@GetMapping("/list")
//	public Page<LikedSongs> list(Pageable page) {
//		return null;
//	}
//
//	/**
//	 * 获取
//	 */
//	@GetMapping("/get")
//	public LikedSongs get(Long id) {
//		return likedSongsService.findById(id);
//	}
//
//	/**
//	 * 添加
//	 */
//	@PostMapping("/add")
//	public void add(@RequestBody LikedSongs likedSongs) {
//		likedSongsService.save(likedSongs);
//	}
//
//
//	/**
//	 * 修改
//	 */
//	@PostMapping("/update")
//	public void update(@RequestBody LikedSongs likedSongs) {
//		likedSongsService.save(likedSongs);
//	}
//
//	/**
//	 * 删除
//	 */
//	@PostMapping("/delete")
//	public void delete(Long id) {
//		likedSongsService.deleteById(id);
//	}
//
//}
//
