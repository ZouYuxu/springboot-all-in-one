package com.example.jparest.thread;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class Test1 {
    public static void main(String[] args) {
        Thread t1 = new Thread(){
            @Override
            public void run() {
                log.info("running");
            }
        };
        t1.setName("t1");
        t1.start();
        log.info("main");
    }
}
