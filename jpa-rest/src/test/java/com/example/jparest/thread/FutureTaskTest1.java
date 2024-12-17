package com.example.jparest.thread;

import java.util.concurrent.Callable;
import java.util.concurrent.FutureTask;

public class FutureTaskTest1 {
    public static void main(String[] args) {
        FutureTask<Object> objectFutureTask = new FutureTask<>(new Callable<>() {
            @Override
            public Object call() throws Exception {
                return null;
            }
        });
    }
}
