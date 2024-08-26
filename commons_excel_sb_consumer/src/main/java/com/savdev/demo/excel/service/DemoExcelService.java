package com.savdev.demo.excel.service;

import com.savdev.commons.excel.api.ExcelReaderApi;
import com.savdev.commons.excel.api.ExcelWriterApi;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

import static com.savdev.demo.excel.config.ExcelConstants.EXCEL_SHEET_NAME;
import static com.savdev.demo.excel.config.ExcelConstants.EXCEL_TEMPLATE;
import static com.savdev.demo.excel.config.ExcelConstants.HEADER_LINE_NUMBER;

@Service
public class DemoExcelService {

  @Autowired
  private ExcelReaderApi excelReaderApi;

  @Autowired
  private ExcelWriterApi excelWriterApi;

  public List<Map<String, Object>> extractLines(InputStream inputStream) {
    return excelReaderApi.linesStream(EXCEL_SHEET_NAME, inputStream)
      .filter(excelLine -> excelLine.getExcelLineNumber() > HEADER_LINE_NUMBER)
      .map(excelReaderApi::transformer)
      .toList();
  }

  public void downloadAsFile(List<Map<String, Object>> lines, OutputStream outputStream) {
    try (InputStream template = getClass().getClassLoader().getResourceAsStream(EXCEL_TEMPLATE);) {
      excelWriterApi.writeExcelTemplate(
        template,
        EXCEL_SHEET_NAME,
        HEADER_LINE_NUMBER + 1,
        lines,
        outputStream
      );
    } catch (IOException e) {
      throw new IllegalStateException(e);
    }
  }
}
