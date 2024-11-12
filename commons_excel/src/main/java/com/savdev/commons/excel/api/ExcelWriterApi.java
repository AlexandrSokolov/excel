package com.savdev.commons.excel.api;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

public interface ExcelWriterApi {

  InputStream writeExcelAsInputStream(
    InputStream excelTemplate,
    String sheetName,
    int firstLineNumber,
    List<Map<String,Object>> rows);

  InputStream writeExcelAsInputStream(
    String sheetName,
    int firstLineNumber,
    List<Map<String,Object>> rows);

  void writeExcelTemplate(
    InputStream excelTemplate,
    String sheetName,
    int firstLineNumber,
    List<Map<String,Object>> rows,
    OutputStream outputStream);

  void writeExcel(
    String sheetName,
    int firstLineNumber,
    List<Map<String,Object>> rows,
    OutputStream outputStream);
}
