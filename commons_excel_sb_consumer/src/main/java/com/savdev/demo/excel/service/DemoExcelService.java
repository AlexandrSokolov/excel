package com.savdev.demo.excel.service;

import com.savdev.commons.excel.api.ExcelReaderApi;
import com.savdev.commons.excel.api.ExcelWriterApi;
import com.savdev.demo.excel.service.validation.collector.ValidationCollector;
import com.savdev.demo.excel.service.validators.ColumnAValidator;
import com.savdev.demo.excel.service.validators.ColumnEValidator;
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

  @Autowired
  private ColumnAValidator columnAValidator;

  @Autowired
  private ColumnEValidator columnEValidator;

  public List<Map<String, Object>> extractLines(InputStream inputStream) {
    var validatedItems = excelReaderApi.linesStream(EXCEL_SHEET_NAME, inputStream)
      .filter(excelLine -> excelLine.getExcelLineNumber() > HEADER_LINE_NUMBER)
      .collect(new ValidationCollector<>(
        List.of(columnAValidator, columnEValidator),
        excelReaderApi::transformer,
        line -> "Line number: " + line.getExcelLineNumber()
      ));
    if (validatedItems.invalidItems().isEmpty()) {
      return validatedItems.validItems();
    } else {
      return validatedItems.invalidItems().stream()
        .map(invalidItem ->
          Map.of(
            "line", invalidItem.itemRepresentation(),
            "errors", invalidItem.validationErrors()))
        .toList();
    }
  }

  public void downloadAsFile(List<Map<String, Object>> lines, OutputStream outputStream) {
    try (InputStream template = getClass().getClassLoader().getResourceAsStream(EXCEL_TEMPLATE)) {
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
