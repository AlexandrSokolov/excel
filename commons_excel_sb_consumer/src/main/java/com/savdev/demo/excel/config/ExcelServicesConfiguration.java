package com.savdev.demo.excel.config;

import com.savdev.commons.excel.api.ExcelReaderApi;
import com.savdev.commons.excel.api.ExcelWriterApi;
import com.savdev.commons.excel.service.ExcelReaderService;
import com.savdev.commons.excel.service.ExcelWriterService;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import java.util.Map;

@Configuration
public class ExcelServicesConfiguration {

  //attribute names for C and D are not set
  public final static Map<String, String> configuration = Map.of(
      "A", "colA",
      "B", "colB",
      "E", "money",
      "F", "persentage",
      "G", "someDate",
      "H", "time",
      "I", "number");

  @Bean
  public ExcelReaderApi excelReaderApi() {
    return ExcelReaderService.instance(configuration);
  }

  @Bean
  public ExcelWriterApi excelWriterApi() {
    return new ExcelWriterService();
  }
}
