package com.savdev.demo.excel.config;

import com.savdev.commons.excel.api.ExcelReaderApi;
import com.savdev.commons.excel.api.ExcelWriterApi;
import com.savdev.commons.excel.service.ExcelReaderService;
import com.savdev.commons.excel.service.ExcelWriterService;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import java.util.Map;

import static com.savdev.demo.excel.config.ExcelConstants.COLUMMN_A_ATTRIBUTE_NAME;
import static com.savdev.demo.excel.config.ExcelConstants.COLUMMN_E_ATTRIBUTE_NAME;

@Configuration
public class ExcelServicesConfiguration {

  //attribute names for C and D are not set
  public final static Map<String, String> configuration = Map.of(
      "A", COLUMMN_A_ATTRIBUTE_NAME,
      "B", "colB",
      "E", COLUMMN_E_ATTRIBUTE_NAME,
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
