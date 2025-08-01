package com.example.demo;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class OfficeProcessorApplication {

    public static void main(String[] args) {
        SpringApplication.run(OfficeProcessorApplication.class, args);
        System.out.println("🚀 Office Processor 启动成功！");
        System.out.println("📄 访问: http://localhost:8080");
    }
}