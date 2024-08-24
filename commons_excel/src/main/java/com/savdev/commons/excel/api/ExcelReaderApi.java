package com.savdev.commons.excel.api;

import com.savdev.commons.excel.dto.ExcelLine;

import java.io.InputStream;
import java.util.Map;
import java.util.function.Supplier;
import java.util.stream.Stream;

public interface ExcelReaderApi {

  /**
   * Excel lines as stream for the 1st Excel sheet
   *
   * @param inputStream
   *
   * @return stream of Excel lines
   */
  Stream<ExcelLine> linesStream(InputStream inputStream);

  /**
   * Excel lines as stream for the specific Excel sheet
   *
   * @param excelSheetName Excel sheet name
   * @param inputStream
   *
   * @return stream of Excel lines
   */
  Stream<ExcelLine> linesStream(
    String excelSheetName,
    InputStream inputStream);

  /**
   * Transforms `ExcelLine` into `attribute` -> `field value` mapping
   *
   * Works only with Excel Readers configured with column -> attribute mapping
   *
   * @param excelLine
   * @return
   */
  Map<String, Object> transformer(ExcelLine excelLine);

  /**
   * Transforms `attribute` -> `field value` mapping into custom DTO object
   *
   * Java Records are not supported. You can transform manually into Record
   *
   * @param supplier - DTO object supplier
   * @param attribute2Value - `named attribute` -> `column value` mapping
   *
   * @return
   */
  <R> R transformer(Supplier<R> supplier, Map<String, Object> attribute2Value);
}
