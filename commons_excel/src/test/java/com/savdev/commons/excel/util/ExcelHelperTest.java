package com.savdev.commons.excel.util;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

public class ExcelHelperTest {

  @Test
  public void testExcelColumns() {
    var columnNames = ExcelHelper.excelColumns(50);
    Assertions.assertEquals("A", columnNames.getFirst());
    Assertions.assertEquals("AX", columnNames.getLast());
  }
}
