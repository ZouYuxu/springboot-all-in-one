package com.example.jparest.model.entity;

import java.util.Date;
import java.io.Serializable;

/**
 * (LikedSongs)实体类
 *
 * @author makejava
 * @since 2023-11-12 10:10:51
 */
public class LikedSongs implements Serializable {
    private static final long serialVersionUID = 469399552958915936L;
    
    private Long userId;
    
    private Date createdAt;
    
    private Long songId;


    public Long getUserId() {
        return userId;
    }

    public void setUserId(Long userId) {
        this.userId = userId;
    }

    public Date getCreatedAt() {
        return createdAt;
    }

    public void setCreatedAt(Date createdAt) {
        this.createdAt = createdAt;
    }

    public Long getSongId() {
        return songId;
    }

    public void setSongId(Long songId) {
        this.songId = songId;
    }

}

