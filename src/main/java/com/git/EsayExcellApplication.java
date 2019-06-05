package com.git;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
@MapperScan("com.git.poi.xlsx")
public class EsayExcellApplication {

    public static void main(String[] args) {
        SpringApplication.run(EsayExcellApplication.class, args);
    }

}
