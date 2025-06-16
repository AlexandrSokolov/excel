package com.savdev.commons.excel.service;

import com.savdev.commons.excel.api.ExcelWriterApi;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.util.*;
import java.util.stream.Stream;

public class ExcelWriterService implements ExcelWriterApi {

  @Override
  public InputStream writeExcelAsInputStream(
    final InputStream excelTemplate,
    final String sheetName,
    final int firstLineNumber,
    final List<Map<String, Object>> rows) {
    try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
      writeExcelTemplate(
        excelTemplate,
        sheetName,
        1, //no template, no header
        rows,
        outputStream);
      return new ByteArrayInputStream(outputStream.toByteArray());
    } catch (IOException e) {
      throw new IllegalStateException(e);
    }
  }

  @Override
  public InputStream writeExcelAsInputStream(String sheetName, int firstLineNumber, List<Map<String, Object>> rows) {
    try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
      writeExcel(
        sheetName,
        1, //no template, no header
        rows,
        outputStream);
      return new ByteArrayInputStream(outputStream.toByteArray());
    } catch (IOException e) {
      throw new IllegalStateException(e);
    }
  }

  @Override
  public InputStream writeExcelAsInputStream(Map<String, Stream<Map.Entry<Integer, Map<String,Object>>>> linesPerSheet) {
    try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
         var workbook = new XSSFWorkbook()) {

      linesPerSheet.forEach((sheetName, lines) -> {
        var sheet = Optional.ofNullable(workbook.getSheet(sheetName))
          .orElse(workbook.createSheet(sheetName));

        var columnNames = new HashSet<String>();
        lines.forEach((entry) -> {
          writeExcelRow(
            sheet,
            //lines in excel start with 1, but in api with 0 index
            entry.getKey() - 1,
            entry.getValue());
          columnNames.addAll(
            entry.getValue().keySet());
        });

        columnNames.forEach(excelColumn -> {
            sheet.autoSizeColumn(
              CellReference.convertColStringToIndex(excelColumn));
          });
      });

      workbook.write(outputStream);

      return new ByteArrayInputStream(outputStream.toByteArray());
    } catch (IOException e) {
      throw new IllegalStateException("Could not write, error: " + e.getMessage(), e);
    }
  }

  @Override
  public void writeExcelTemplate(
    final InputStream excelTemplate,
    final String sheetName,
    final int firstLineNumber,
    final List<Map<String,Object>> rows,
    final OutputStream outputStream
    ) {
    try {
      XSSFWorkbook workbook = new XSSFWorkbook(excelTemplate);

      XSSFSheet sheet = Optional.ofNullable(sheetName)
        .map(name -> workbook.getSheet(sheetName))
        .orElse(workbook.getSheetAt(0));

      writeExcelSheet(
        workbook,
        sheet,
        firstLineNumber,
        rows,
        outputStream);

    } catch (IOException e) {
      throw new RuntimeException(
        "Could not write using excel template. Error: " + e.getMessage(), e);
    }
  }

  @Override
  public void writeExcel(
    final String sheetName,
    final int firstLineNumber,
    final List<Map<String,Object>> rows,
    final OutputStream outputStream) {

    XSSFWorkbook workbook = new XSSFWorkbook();

    XSSFSheet sheet = Optional.ofNullable(sheetName)
      .map(name -> workbook.createSheet(sheetName))
      .orElse(workbook.createSheet());

    writeExcelSheet(
      workbook,
      sheet,
      firstLineNumber,
      rows,
      outputStream);

  }

  private void writeExcelSheet(
    final Workbook workbook,
    final Sheet sheet,
    final int firstLineNumber,
    final List<Map<String,Object>> rows,
    final OutputStream outputStream) {
    final int [] lineNumber = { firstLineNumber != 0 ? firstLineNumber-1 : firstLineNumber };

    rows.forEach(rowMap -> writeExcelRow(sheet, lineNumber[0]++, rowMap));

    rows.getFirst().keySet().forEach(excelColumn -> {
      sheet.autoSizeColumn(
        CellReference.convertColStringToIndex(excelColumn));
    });

    try {
      workbook.write(outputStream);
    } catch (IOException e) {
      throw new RuntimeException("Could not write, error: " + e.getMessage(), e);
    }
  }

  private void writeExcelRow(final Sheet sheet, Integer lineNumber, final Map<String,Object> excelRow) {
    var row = sheet.createRow(lineNumber);
    excelRow.forEach((excelColumn, excelColValue) ->
      createCell(row, excelColumn, excelColValue));
  }

  private void createCell(
    final Row row,
    final String excelColumn,
    Object cellValue) {
    final int index = CellReference.convertColStringToIndex(excelColumn);
    if (cellValue != null) {
      if (cellValue instanceof Double
        || cellValue instanceof Integer) {
        Cell cell = row.createCell(index, CellType.NUMERIC);
        cell.setCellValue((double) cellValue);
      } else if (cellValue instanceof BigDecimal) {
        Cell cell = row.createCell(index, CellType.NUMERIC);
        cell.setCellValue(((BigDecimal)cellValue).doubleValue());
      } else if (cellValue instanceof Boolean) {
        Cell cell = row.createCell(index, CellType.BOOLEAN);
        cell.setCellValue(((Boolean) cellValue));
      } else {
        Cell cell = row.createCell(index, CellType.STRING);
        cell.setCellValue(cellValue.toString());
      }
    } else {
      row.createCell(index, CellType.BLANK);
    }
  }

}
