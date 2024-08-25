package com.savdev.commons.excel.service;

import com.savdev.commons.excel.api.ExcelWriterApi;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.List;
import java.util.Map;
import java.util.Optional;

public class ExcelWriterService implements ExcelWriterApi {

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

    rows.forEach(rowMap -> {
      Row row = sheet.createRow(lineNumber[0]++);
      rowMap.forEach((excelColumn, excelColValue) ->
        createCell(row, excelColumn, excelColValue));
    });

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
