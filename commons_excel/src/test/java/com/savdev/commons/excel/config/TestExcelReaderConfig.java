package com.savdev.commons.excel.config;

import com.savdev.commons.excel.api.ExcelReaderApi;
import com.savdev.commons.excel.service.ExcelReaderService;
import org.springframework.boot.test.context.TestConfiguration;
import org.springframework.context.annotation.Bean;

import java.util.Map;

@TestConfiguration
public class TestExcelReaderConfig {

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
}
