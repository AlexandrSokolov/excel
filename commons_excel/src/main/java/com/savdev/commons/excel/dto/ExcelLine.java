package com.savdev.commons.excel.dto;

import java.util.AbstractMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.StringJoiner;

public class ExcelLine {

  private final int excelLineNumber;

  private final Map<String, Map.Entry<String,  Object>> columnName2Attribute2Value = new LinkedHashMap<>();

  private ExcelLine(int excelLineNumber) {
    this.excelLineNumber = excelLineNumber;
  }

  public static ExcelLine instance(int excelLineNumber) {
    return new ExcelLine(excelLineNumber);
  }

  public ExcelLine put(String excelColName, String attributeName, Object value) {
    this.columnName2Attribute2Value.put(excelColName, new AbstractMap.SimpleEntry<>(attributeName, value));
    return this;
  }

  public Optional<Object> valueByColumn(final String columnName) {
    return Optional.ofNullable(columnName2Attribute2Value.get(columnName))
      .map(Map.Entry::getValue);
  }

  public Optional<Object> valueByAttributeName(final String attributeName) {
    return Optional.ofNullable(columnName2Attribute2Value.values().stream()
      .collect(LinkedHashMap::new,
        (m, v) -> m.put(v.getKey(), v.getValue()),
        Map::putAll)
      .get(attributeName));
  }

  public Map<String, Map.Entry<String,  Object>> getColumnName2Attribute2Value() {
    return columnName2Attribute2Value;
  }

  public int getExcelLineNumber() {
    return excelLineNumber;
  }

  @Override
  public boolean equals(Object o) {
    if (this == o) return true;
    if (o == null || getClass() != o.getClass()) return false;
    ExcelLine excelLine = (ExcelLine) o;
    return excelLineNumber == excelLine.excelLineNumber
      && Objects.equals(columnName2Attribute2Value, excelLine.columnName2Attribute2Value);
  }

  @Override
  public int hashCode() {
    return Objects.hash(excelLineNumber, columnName2Attribute2Value);
  }

  @Override
  public String toString() {
    return new StringJoiner(", ", ExcelLine.class.getSimpleName() + "[", "]")
      .add("excelLineNumber=" + excelLineNumber)
      .add("columnName2Attribute2Value=" + columnName2Attribute2Value)
      .toString();
  }
}
