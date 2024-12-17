package com.example.jparest.thread;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class Test2 {
    public static void main(String[] args) {
        Runnable runnable = () -> log.info("runnable");
        Thread t1 = new Thread(runnable, "t1");
        t1.start();

        new Thread(() -> log.info("runnable2")).start();
    }
}
