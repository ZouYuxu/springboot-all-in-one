package org.zyx;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.cloud.client.discovery.EnableDiscoveryClient;

@SpringBootApplication
@EnableDiscoveryClient
public class TripGateWayApplication {
    public static void main(String[] args) {
        SpringApplication.run(TripGateWayApplication.class, args);
    }
}
