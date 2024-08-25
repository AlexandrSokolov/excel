package com.savdev.commons.excel.api;

import com.savdev.commons.excel.CommonsExcelDiConfig;
import com.savdev.commons.excel.TestPathUtils;
import com.savdev.commons.excel.config.TestExcelReaderConfig;
import com.savdev.commons.excel.dto.ExcelLine;
import com.savdev.commons.excel.service.ExcelReaderService;
import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.ContextConfiguration;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import static com.savdev.commons.excel.service.ExcelReaderService.CANNOT_CONVERT_WITHOUT_ATTRIBUTES_MAPPING;
import static org.junit.jupiter.api.Assertions.assertThrows;

@SpringBootTest
@ContextConfiguration(classes = {
  CommonsExcelDiConfig.class,
  TestExcelReaderConfig.class})
public class ExcelReaderApiTest {

  private static final String EXCEL_SHEET_NAME = "Test 1";
  //in this line, line #3 defines attributes/fields
  private static final Integer HEADER_LINE_NUMBER = 3;
  private static final String EXCEL_FILE = "9Columns2Rows.xlsx";

  private static final Logger logger = LogManager.getLogger(ExcelReaderApiTest.class.getName());

  private final TestPathUtils testPathUtils = new TestPathUtils();

  @Autowired
  private ExcelReaderApi excelReaderApi;

  /**
   * Test injected instance of `ExcelReaderApi`
   *
   *  The key-field configuration of Excel reader is provided in `TestExcelReaderConfig`
   *
   * @throws IOException
   */
  @Test
  public void testConfiguredReader() throws IOException {
    try (InputStream stream = testPathUtils.testInputStream(EXCEL_FILE)) {
      Assertions.assertEquals(
        expectedLines(),
        excelReaderApi.linesStream(EXCEL_SHEET_NAME, stream).toList());
    }
  }

  /**
   * The key-field configuration of Excel reader is taken from the header line in Excel
   *
   * @throws IOException
   */
  @Test
  public void testConfiguredFromLineReader() throws IOException {
    var requiredAttributes = Set.of("colA", "colB");
    try (InputStream stream = testPathUtils.testInputStream(EXCEL_FILE)) {
      var lines = ExcelReaderService.instance(
        HEADER_LINE_NUMBER,
        column2Attribute -> {
          var notExisting = requiredAttributes.stream()
            .filter(a -> !column2Attribute.containsValue(a))
            .toList();
          if (!notExisting.isEmpty()) {
            throw new IllegalStateException("Required, but not existing attributes: '"
              + notExisting + "' in Excel line #" + HEADER_LINE_NUMBER);
          }
        })
        .linesStream(EXCEL_SHEET_NAME, stream)
        .toList();

      var line4 = lines.getFirst();
      Assertions.assertEquals(4, line4.getExcelLineNumber());
      //assert column attribute, extracted from the header line
      Assertions.assertEquals("colA", line4.getColumnName2Attribute2Value().get("A").getKey());
      //assert column value, extracted from the line
      Assertions.assertEquals("line_4a", line4.getColumnName2Attribute2Value().get("A").getValue());

      //assert column attribute, extracted from the header line
      Assertions.assertEquals("money", line4.getColumnName2Attribute2Value().get("E").getKey());
      //assert column value, extracted from the line
      Assertions.assertEquals(new BigDecimal("32.5"), line4.getColumnName2Attribute2Value().get("E").getValue());

      var line5 = lines.getLast();
      Assertions.assertEquals(5, line5.getExcelLineNumber());
      //assert column attribute, extracted from the header line
      Assertions.assertEquals("colA", line5.getColumnName2Attribute2Value().get("A").getKey());
      //assert column value, extracted from the line
      Assertions.assertEquals("line_5a", line5.getColumnName2Attribute2Value().get("A").getValue());

      //assert column attribute, extracted from the header line
      Assertions.assertEquals("money", line5.getColumnName2Attribute2Value().get("E").getKey());
      //assert column value, extracted from the line
      Assertions.assertEquals(new BigDecimal("44.8"), line5.getColumnName2Attribute2Value().get("E").getValue());
    }
  }

  @Test
  public void testConfiguredFromLineHeaderValidationFailed() {
    final var requiredAttributes = Set.of("colA", "colB", "colC", "colD");
    Exception e = Assertions.assertThrows(
      IllegalStateException.class,
      () -> {
        try (InputStream stream = testPathUtils.testInputStream(EXCEL_FILE)) {
          ExcelReaderService.instance(
              HEADER_LINE_NUMBER,
              column2Attribute -> {
                var notExisting = requiredAttributes.stream()
                  .filter(a -> !column2Attribute.containsValue(a))
                  .toList();
                if (!notExisting.isEmpty()) {
                  throw new IllegalStateException("Required, but not existing attributes: '"
                    + notExisting + "' in Excel line #" + HEADER_LINE_NUMBER);
                }
              })
            .linesStream(EXCEL_SHEET_NAME, stream)
            .toList();
        }
      });
    Assertions.assertEquals(
      "Required, but not existing attributes: '[colC, colD]' in Excel line #3",
      e.getMessage()
    );
  }

  /**
   * No key-field mapping configuration
   *
   * @throws IOException
   */
  @Test
  public void testNonConfiguredReader() throws IOException {

    var notConifiguredExcelReaderApi = ExcelReaderService.instance();

    try (InputStream stream = testPathUtils.testInputStream(EXCEL_FILE)) {
      //note: each line contains empty null value as `field`
      var lines = notConifiguredExcelReaderApi.linesStream(stream)
        .peek(line -> logger.info("column B value: " +
          line.getColumnName2Attribute2Value()
            .get("B")
            .getValue()))
        .toList();
      //cannot be compared with result of `expectedLines()`
      Assertions.assertEquals(expectedLines().size(), lines.size());
    }
  }

  /**
   * Example of
   * - filtering (only line numbers > 2) and
   * - validation (valid are lines that have non-empty value in `D` column)
   *    *
   * @throws IOException
   */
  @Test
  public void testFilteringAndValidation() throws IOException {
    try (InputStream stream = testPathUtils.testInputStream(EXCEL_FILE)) {

      Map<Boolean, List<ExcelLine>> validAndInvalid = excelReaderApi.linesStream(EXCEL_SHEET_NAME, stream)
        .filter(excelLine -> excelLine.getExcelLineNumber() > HEADER_LINE_NUMBER)
        .collect(Collectors.partitioningBy(excelLine -> excelLine.valueByColumn("D")
          .map(value -> (String) value)
          .map(StringUtils::isNoneEmpty)
          .isPresent()));
      Assertions.assertEquals(2, validAndInvalid.size());
      //line #5 is valid, it has non-empty value in the `D` column
      Assertions.assertEquals(1, validAndInvalid.get(true).size());
      //line #3 and #4 have empty value in the `D` column
      Assertions.assertEquals(1, validAndInvalid.get(false).size());
    }
  }

  /**
   * Transform `ExcelLine` type into `field` -> value mapping
   *
   * Works only with configured column-field readers
   *
   * @throws IOException
   */
  @Test
  public void testTransformer() throws IOException {
    try (InputStream stream = testPathUtils.testInputStream(EXCEL_FILE)) {
      List<Map<String, Object>> result = excelReaderApi.linesStream(EXCEL_SHEET_NAME, stream)
        .filter(excelLine -> excelLine.getExcelLineNumber() > HEADER_LINE_NUMBER)
        .map(excelReaderApi::transformer)
        .toList();
      //check that all "fields" are extracted according to the configuration
      Assertions.assertTrue(result.stream().allMatch(
        line -> TestExcelReaderConfig.configuration
          .values().stream()
          .allMatch(line::containsKey)));
      Assertions.assertEquals(2, result.size());
    }
  }

  /**
   * It cannot transform `ExcelLine` type into `field` -> value mapping with no column -> attribute configuration
   *
   * @throws IOException
   */
  @Test
  public void testNotConfiguredReaderTransformer() throws IOException {
    try (InputStream stream = testPathUtils.testInputStream(EXCEL_FILE)) {

      Exception exception = assertThrows(
        IllegalStateException.class,
        () -> ExcelReaderService.instance()
          .linesStream(EXCEL_SHEET_NAME, stream)
          .filter(excelLine -> excelLine.getExcelLineNumber() > HEADER_LINE_NUMBER)
          .map(excelReaderApi::transformer)
          .toList());

      Assertions.assertTrue(
        exception.getMessage()
          .startsWith(CANNOT_CONVERT_WITHOUT_ATTRIBUTES_MAPPING));
    }
  }

  @Test
  public void testTransformer2Dto() throws IOException {
    try (InputStream stream = testPathUtils.testInputStream(EXCEL_FILE)) {
      List<TestDto> result = excelReaderApi.linesStream(EXCEL_SHEET_NAME, stream)
        .filter(excelLine -> excelLine.getExcelLineNumber() > HEADER_LINE_NUMBER)
        .map(excelReaderApi::transformer)
        .map(attribute2Value -> excelReaderApi.transformer(
          TestDto::new,
          attribute2Value))
        .toList();
      Assertions.assertEquals(2, result.size());
    }
  }

  private List<ExcelLine> expectedLines() {
    return List.of(
      line1(),
      //line 2 has null values
      line3(),
      line4(),
      line5());
  }

  private ExcelLine line1() {
    return ExcelLine.instance(1).put("B", "colB", "test");
  }

  private ExcelLine line3() {
    return ExcelLine.instance(3)
      .put("A", "colA", "colA")
      .put("B", "colB", "colB")
      .put("E", "money", "money")
      .put("F", "persentage", "persentage")
      .put("G", "someDate", "someDate")
      .put("H", "time", "time")
      .put("I", "number", "number");
  }

  private ExcelLine line4() {
    return ExcelLine.instance(4)
      .put("A", "colA", "line_4a")
      .put("B", "colB", "line_4b")
      .put("C", null, "line_4_c")
      .put("E", "money", BigDecimal.valueOf(32.5))
      .put("F", "persentage", BigDecimal.valueOf(23.5))
      .put("G", "someDate",
        LocalDateTime.parse("2021-08-06T00:00", DateTimeFormatter.ISO_LOCAL_DATE_TIME))
      .put("H", "time", LocalTime.parse("16:12:37.239"))
      .put("I", "number", BigDecimal.valueOf(23.4));
  }

  private ExcelLine line5() {
    return ExcelLine.instance(5)
      .put("A", "colA", "line_5a")
      .put("B", "colB", "line_5b")
      .put("D", null, "line_5D")
      .put("E", "money", BigDecimal.valueOf(44.80))
      .put("F", "persentage", BigDecimal.valueOf(47.89))
      .put("G", "someDate",
        LocalDateTime.parse("23-08-2021T00:00", DateTimeFormatter.ofPattern("dd-MM-yyyy'T'HH:mm")))
      .put("H", "time", LocalTime.parse("16:12:44.723"))
      .put("I", "number", BigDecimal.valueOf(43.884));
  }
}
