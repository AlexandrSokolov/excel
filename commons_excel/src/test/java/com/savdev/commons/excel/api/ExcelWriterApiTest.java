package com.savdev.commons.excel.api;

import com.savdev.commons.excel.CommonsExcelDiConfig;
import com.savdev.commons.excel.TestPathUtils;
import com.savdev.commons.excel.service.ExcelWriterService;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.ContextConfiguration;

import java.io.*;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.AbstractMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

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

  /**
   * Write Excel and return it as InputStream
   *
   * @throws IOException
   */
  @Test
  public void testWriteExcelAsInputStreamNoTemplate() throws IOException {
    try (InputStream inputStream = excelWriterApi.writeExcelAsInputStream(
      EXCEL_SHEET_NAME,
      1, //no template, no header
      excelLines())) {
      var targetFile = tempFile();
      Files.copy(
        inputStream,
        targetFile.toPath(),
        StandardCopyOption.REPLACE_EXISTING);
      System.out.println(targetFile.getAbsolutePath());
    }
  }

  /**
   * Write Excel and return it as InputStream
   *
   * @throws IOException
   */
  @Test
  public void testWriteExcelAsInputStreamWithTemplate() throws IOException {
    try (InputStream template = testPathUtils.testInputStream(TEMPLATE);
      InputStream inputStream = excelWriterApi.writeExcelAsInputStream(
        template,
      EXCEL_SHEET_NAME,
        HEADER_LINE_NUMBER + 1,
      excelLines())) {
      var targetFile = tempFile();
      Files.copy(
        inputStream,
        targetFile.toPath(),
        StandardCopyOption.REPLACE_EXISTING);
      System.out.println(targetFile.getAbsolutePath());
    }
  }

  @Test
  public void testWriteExcelAsInputStreamDifferentSheets() throws IOException {

    //use LinkedHashMap to make sure sheets are defined in the expected order
    var orderedMap = new LinkedHashMap<String, Stream<Map.Entry<Integer, Map<String, Object>>>>();
    orderedMap.put(
      EXCEL_SHEET_NAME,
      Stream.of(
        excelLine( 2, Map.of("A", "ABC", "B", new BigDecimal("47.5"))),
        excelLine( 3, Map.of("A", "BC", "C", new BigDecimal("777.5")))
      ));
    orderedMap.put(
      EXCEL_SHEET_NAME + " - another",
      Stream.of(
        excelLine( 1, Map.of("A", "ABC - another", "B", new BigDecimal("555.5"))),
        excelLine( 5, Map.of("A", "BC - next", "C", new BigDecimal("999.5")))));

    try (InputStream inputStream = excelWriterApi.writeExcelAsInputStream(orderedMap)) {
      var targetFile = targetFile(OUTPUT_EXCEL_FILE);
      Files.copy(
        inputStream,
        targetFile.toPath(),
        StandardCopyOption.REPLACE_EXISTING);
      System.out.println(targetFile.getAbsolutePath());
    }
  }

  private Map.Entry<Integer, Map<String, Object>> excelLine(
    Integer lineNumber,
    Map<String, Object> excelRow) {
    return new AbstractMap.SimpleEntry<>(
      lineNumber,
      excelRow);
  }

  /**
   * The content of Excel is put into the InputStream
   *  and then is saved into the temp file, to check the content
   *
   * @throws IOException
   */
  @Test
  public void testWriteExcelIntoInputStream2() throws IOException {
    try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
      excelWriterApi.writeExcel(
        EXCEL_SHEET_NAME,
        1, //no template, no header
        excelLines(),
        outputStream);
      try (InputStream inputStream = new ByteArrayInputStream(outputStream.toByteArray())) {
        var targetFile = tempFile();
        Files.copy(
          inputStream,
          targetFile.toPath(),
          StandardCopyOption.REPLACE_EXISTING);
        System.out.println(targetFile.getAbsolutePath());
      }
    }
  }

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

  private File targetFile(String fileName){
    String projectDir = System.getProperty("user.dir");
    String targetPath = projectDir + "/target";
    File targetDir = new File(targetPath);

    if (!targetDir.exists()) {
      targetDir.mkdirs();
    }

    return Paths.get(targetPath, fileName).toFile();
  }

  private List<Map<String, Object>> excelLines() {
    return List.of(
      Map.of("A", "ABC", "E", new BigDecimal("45.64")),
      Map.of("A", "DEF", "E", new BigDecimal("275.45")));
  }
}
