package cn.seppel.smarttool;


import cn.seppel.smarttool.service.ExcelTransformerService;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.WebApplicationType;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;

@SpringBootApplication
public class SmartToolApplication implements CommandLineRunner {

    private final ExcelTransformerService transformerService;

    public SmartToolApplication(ExcelTransformerService transformerService) {
        this.transformerService = transformerService;
    }

    public static void main(String[] args) {
        SpringApplication app = new SpringApplication(SmartToolApplication.class);
        app.setWebApplicationType(WebApplicationType.NONE);
        app.run(args);
    }

    @Override
    public void run(String... args) throws Exception {
        String inputPath = "input\\";
        String outputPath = "total invoices.xlsx";
        transformerService.transform(inputPath, outputPath);
        System.out.println("输出结果文件：" + outputPath);
        try {
            Thread.sleep(1000);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
        }
    }
}