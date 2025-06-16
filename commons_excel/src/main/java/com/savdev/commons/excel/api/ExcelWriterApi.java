package com.savdev.commons.excel.api;

import com.savdev.commons.excel.dto.ExcelLine;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

public interface ExcelWriterApi {

  /**
   * to write multiple lines into an Excel sheet, with Excel template
   *
   * @return
   */
  InputStream writeExcelAsInputStream(
    InputStream excelTemplate,
    String sheetName,
    int firstLineNumber,
    List<Map<String,Object>> rows);

  /**
   * to write multiple lines into an Excel sheet, without Excel template
   *
   * @return
   */
  InputStream writeExcelAsInputStream(
    String sheetName,
    int firstLineNumber,
    List<Map<String,Object>> rows);

  /**
   * This method allows to write multiple lines into different Excel sheets in a streaming way
   *
   * @param linesPerSheet
   * @return
   */
  InputStream writeExcelAsInputStream(
    Map<String, Stream<Map.Entry<Integer, Map<String,Object>>>> linesPerSheet);

  /**
   * to write multiple lines into an Excel sheet, with Excel template
   *  for rest service, see RestExcelRestApi.downloadFile()
   *
   * @return
   */
  void writeExcelTemplate(
    InputStream excelTemplate,
    String sheetName,
    int firstLineNumber,
    List<Map<String,Object>> rows,
    OutputStream outputStream);

  /**
   * to write multiple lines into an Excel sheet, without Excel template
   *  for rest service, see RestExcelRestApi.downloadFile()
   *
   * @return
   */
  void writeExcel(
    String sheetName,
    int firstLineNumber,
    List<Map<String,Object>> rows,
    OutputStream outputStream);
}
