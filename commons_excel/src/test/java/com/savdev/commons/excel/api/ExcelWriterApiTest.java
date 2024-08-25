package com.savdev.commons.excel.api;

import com.savdev.commons.excel.CommonsExcelDiConfig;
import com.savdev.commons.excel.TestPathUtils;
import com.savdev.commons.excel.service.ExcelWriterService;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.ContextConfiguration;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import static com.savdev.commons.excel.api.ExcelTestConstants.EXCEL_SHEET_NAME;
import static com.savdev.commons.excel.api.ExcelTestConstants.HEADER_LINE_NUMBER;
import static com.savdev.commons.excel.api.ExcelTestConstants.OUTPUT_EXCEL_FILE;
import static com.savdev.commons.excel.api.ExcelTestConstants.TEMPLATE;

@SpringBootTest
@ContextConfiguration(classes = {
  CommonsExcelDiConfig.class})
public class ExcelWriterApiTest {

  private final TestPathUtils testPathUtils = new TestPathUtils();
  private final ExcelWriterApi excelWriterApi = new ExcelWriterService();

  @TempDir
  Path tempTestFolder;

  @Test
  public void testWriteExcelTemplate() throws IOException {
    var tempFile = tempFile();
    try (
      InputStream stream = testPathUtils.testInputStream(TEMPLATE);
      FileOutputStream outputStream = new FileOutputStream(tempFile)) {
      excelWriterApi.writeExcelTemplate(
        stream,
        EXCEL_SHEET_NAME,
        HEADER_LINE_NUMBER + 1,
        excelLines(),
        outputStream);
      System.out.println(tempFile.getAbsolutePath());
    }
  }

  @Test
  public void testWriteExcelTemplateWithoutSheetName() throws IOException {
    var tempFile = tempFile();
    try (
      InputStream stream = testPathUtils.testInputStream(TEMPLATE);
      FileOutputStream outputStream = new FileOutputStream(tempFile)) {
      excelWriterApi.writeExcelTemplate(
        stream,
        null, // do not set sheet name, the 1st sheet is used by default
        HEADER_LINE_NUMBER + 1,
        excelLines(),
        outputStream);
      System.out.println(tempFile.getAbsolutePath());
    }
  }

  @Test
  public void testWriteExcel() throws IOException {
    var tempFile = tempFile();
    try (FileOutputStream outputStream = new FileOutputStream(tempFile)) {
      excelWriterApi.writeExcel(
        EXCEL_SHEET_NAME,
        1, //no template, no header
        excelLines(),
        outputStream);
      System.out.println(tempFile.getAbsolutePath());
    }
  }

  private File tempFile(){
    return new File(tempTestFolder.toFile(), OUTPUT_EXCEL_FILE);
  }

  private List<Map<String, Object>> excelLines() {
    return List.of(
      Map.of("A", "ABC", "E", new BigDecimal("45.64")),
      Map.of("A", "DEF", "E", new BigDecimal("275.45")));
  }
}
