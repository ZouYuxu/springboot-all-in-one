package com.example.jparest.utils;

import lombok.extern.slf4j.Slf4j;
import org.springframework.util.StopWatch;

import java.time.Duration;
import java.time.Instant;
import java.util.function.Function;

@Slf4j
public class Sw {
    private static StopWatch st = new StopWatch();
    private static Instant last;
    private static long total;

    static void start(String s) {
        st.start(s);
    }


    public static void stop() {
        st.stop();

    }

    public static void prettyPrint() {
        log.info(st.prettyPrint());
        log.info("总用时 {} ms", st.getTotalTimeMillis());
    }

    public static void time(Function fun) {
//        fun.apply(this);
    }

    public static void record() {
        last = Instant.now();
    }

    public static void cost() {
        Duration between = Duration.between(last, Instant.now());
        long mi = between.toMillis();
        total += mi;
        log.info(mi + "ms");
    }

    public static void total() {
        log.info(total + "ms");
    }
}
