package com.savdev.commons.excel.util;

import org.apache.poi.ss.util.CellReference;

import java.util.List;
import java.util.stream.IntStream;

public class ExcelHelper {

  public static List<String> excelColumns(Integer columnsNumber) {
    return IntStream.range(0, columnsNumber).mapToObj(CellReference::convertNumToColString).toList();
  }
}
