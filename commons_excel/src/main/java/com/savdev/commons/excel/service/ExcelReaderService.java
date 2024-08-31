package com.savdev.commons.excel.service;

import com.savdev.commons.excel.api.ExcelReaderApi;
import com.savdev.commons.excel.dto.ExcelLine;
import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.UnsupportedFileFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.*;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Supplier;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

import static org.apache.poi.ss.usermodel.CellType.BLANK;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;

public class ExcelReaderService implements ExcelReaderApi {

  private static final Logger logger = LogManager.getLogger(ExcelReaderService.class.getName());

  static final String NOT_XLSX_FILE = "File is not valid XLSX format. Error: '%s'";
  static final String NOT_VALID_FILE = "Could not handle XLSX file. Error: '%s'";
  public static final String CANNOT_CONVERT_WITHOUT_ATTRIBUTES_MAPPING =
    "Cannot convert line to 'attribute -> value' mapping without configured 'column -> attribute' mapping.";

  private final Map<String, String> column2Attribute;

  private final int headerLineNumber;

  private final Consumer<Map<String, String>> headerLineValidator;

  public static ExcelReaderService instance(Map<String, String> column2Attribute) {
    return new ExcelReaderService(column2Attribute, 0, null);
  }

  public static ExcelReaderService instance(
    final int headerLineNumber,
    final Consumer<Map<String, String>> headerLineValidator) {
    return new ExcelReaderService(
      new LinkedHashMap<>(), //not immutable map!
      headerLineNumber,
      headerLineValidator);
  }

  public static ExcelReaderService instance() {
    return new ExcelReaderService(
      Collections.emptyMap(), 0, null);
  }

  private ExcelReaderService(
    final Map<String, String> column2Attribute,
    final int headerLineNumber,
    final Consumer<Map<String, String>> headerLineValidator) {
    this.column2Attribute = column2Attribute;
    this.headerLineNumber = headerLineNumber;
    this.headerLineValidator = headerLineValidator;
  }

  @Override
  public Stream<ExcelLine> linesStream(
    final String excelSheetName,
    final InputStream inputStream) {
    return catchWorkBook(inputStream,
      workBook -> StreamSupport.stream(
        workBook.getSheet(excelSheetName).spliterator(), false)
        .filter(this::anyCellIsNotEmpty)
        .map(this::fromRow)
        .filter(Objects::nonNull));
  }

  public Stream<ExcelLine> linesStream(
    final InputStream inputStream) {
    return catchWorkBook(inputStream,
      workBook -> StreamSupport.stream(
          workBook.getSheetAt(0).spliterator(), false)
        .filter(this::anyCellIsNotEmpty)
        .map(this::fromRow)
        .filter(Objects::nonNull));
  }

  @Override
  public Map<String, Object> transformer(ExcelLine excelLine) {
    if (!excelLine.getColumnName2Attribute2Value().values().isEmpty()
      && excelLine.getColumnName2Attribute2Value().values().stream().map(Map.Entry::getKey).allMatch(Objects::isNull)) {
      throw new IllegalStateException(CANNOT_CONVERT_WITHOUT_ATTRIBUTES_MAPPING + "Line: \n" + excelLine);
    }

    try {
      return excelLine.getColumnName2Attribute2Value().values().stream()
        .filter(entry -> StringUtils.isNoneEmpty(entry.getKey()))
        .collect(LinkedHashMap::new,
          (m, v)-> m.put(v.getKey(), v.getValue()),
          Map::putAll);
    } catch (Exception e) {
      throw new IllegalStateException("Could not transform: " + excelLine + " to map, error: " + e.getMessage());
    }
  }

  /**
   * Note: all attributes, that do not have the corresponding setter are silently skipped
   * @param supplier of POJO. It cannot be java record
   *
   * @param attribute2Value you can get it via `transformer` method
   * @param <R>
   *
   * @return
   */
  @Override
  public <R> R transformer(Supplier<R> supplier, Map<String, Object> attribute2Value) {
    R result = supplier.get(); //todo rewrite via streams
    attribute2Value.entrySet().stream()
      .filter(entry -> StringUtils.isNoneEmpty(entry.getKey()) && entry.getValue() != null)
      .forEach(entry -> {
        try {
          Method setter = result.getClass()
            .getDeclaredMethod("set" + StringUtils.capitalize(entry.getKey()),
            entry.getValue().getClass());
          setter.invoke(result, entry.getValue());
        } catch (NoSuchMethodException e) {
          logger.debug(() -> "No '" + "set" + StringUtils.capitalize(entry.getKey()) + "' setter.");
        } catch (InvocationTargetException | IllegalAccessException e) {
          throw new IllegalStateException(e);
        }
      });
    return result;
  }

  private boolean anyCellIsNotEmpty(Row row) {
    return StreamSupport.stream(row.spliterator(), false)
      .anyMatch(c -> !Objects.equals(BLANK, c.getCellType()));
  }

  private <R> R catchWorkBook(final InputStream inputStream, Function<Workbook, R> consumer) {
    try (Workbook workbook = new XSSFWorkbook(inputStream)) {
      return consumer.apply(workbook);
    } catch (UnsupportedFileFormatException e){
      throw new IllegalStateException(
        String.format(NOT_XLSX_FILE, e.getMessage()));
    } catch (IOException e) {
      throw new IllegalStateException(
        String.format(NOT_VALID_FILE, e.getMessage()));
    }
  }

  private ExcelLine fromRow(Row row) {
    int currentLineNumber = row.getRowNum() + 1;
    if (headerLineNumber > 0) {
      if (currentLineNumber == headerLineNumber) {
        extractMappingFromHeaderLine(row);
        headerLineValidator.accept(column2Attribute);
      }

      //skip all the lines before the header line, including the header line
      if (currentLineNumber <= headerLineNumber) {
        return null;
      }
    }

    ExcelLine line = ExcelLine.instance(currentLineNumber);
    StreamSupport.stream(row.spliterator(), false)
      .filter(this::isExpectedCellType)
      .forEach(cell -> {
        String columnName = CellReference.convertNumToColString(cell.getColumnIndex());
        line.put(columnName, column2Attribute.get(columnName), extractValue(cell));
      });
    return line;
  }

  private void extractMappingFromHeaderLine(final Row row) {
    ExcelLine line = ExcelLine.instance(row.getRowNum() + 1);
    if (!column2Attribute.isEmpty() && line.getExcelLineNumber() != headerLineNumber) {
      throw new IllegalStateException("Wrong API usage. " +
        "Mapping can be extracted only from the header line = '" + headerLineNumber + "'" +
        "No mapping configuration must be predefined");
    }
    StreamSupport.stream(row.spliterator(), false)
      .filter(this::isExpectedCellType)
      .forEach(cell -> {
        String columnName = CellReference.convertNumToColString(cell.getColumnIndex());
        line.put(columnName, null, extractValue(cell));
      });

    column2Attribute.putAll(
      line.getColumnName2Attribute2Value().entrySet().stream()
        .filter(entry ->
          entry.getValue() != null
            && entry.getValue().getValue() != null
            && StringUtils.isNoneEmpty(entry.getValue().getValue().toString()))
        .map(entry -> new AbstractMap.SimpleEntry<>(entry.getKey(), entry.getValue().getValue().toString()))
        .collect(LinkedHashMap::new,
          (m, v)-> m.put(v.getKey(), v.getValue()),
          Map::putAll));
  }

  private boolean isExpectedCellType(Cell cell) {
    switch (cell.getCellType()) {
      case BLANK:
      case STRING:
      case NUMERIC:
      case BOOLEAN: return true;
      default: return false;
    }
  }

  private Object extractValue(Cell cell){
    if (cell == null) {
      return null;
    }

    switch (cell.getCellType()) {
      case BLANK: return null;
      case STRING: return cell.getStringCellValue();
      case NUMERIC: return fromNumericCell(cell);
      case BOOLEAN: return cell.getBooleanCellValue();
      default: throw new IllegalStateException("Unsupported cell type");
    }
  }

  private Object fromNumericCell(Cell cell){
    if (NUMERIC != cell.getCellType()) {
      throw new IllegalStateException("Wrong API usage. Only for numeric cells could be used.");
    }
    if (cell.getCellStyle().getDataFormatString().contains("%")) {
      return BigDecimal.valueOf(cell.getNumericCellValue() * 100);
    } else if (DateUtil.isCellDateFormatted(cell)) {
      if (cell.getCellStyle().getDataFormatString().startsWith("hh")) {
        return cell.getLocalDateTimeCellValue().toLocalTime();
      } else {
        return cell.getLocalDateTimeCellValue();
      }
    } else {
      return BigDecimal.valueOf(cell.getNumericCellValue());
    }
  }
}
