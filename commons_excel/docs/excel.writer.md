#### The purpose of this Excel Writer is - to hide excel-library specific logic and expose only standard java api types

**Note:** this Excel Writer uses the `column letter -> value` mapping, but not the `attribute -> value`

### How to use Excel Writer:

- [Write into Excel template](#write-into-excel-template)
- [Write into Excel without template](#write-into-excel-without-template)
- [Writing without having Excel sheet name defined](#writing-without-having-excel-sheet-name-defined)

### Write into Excel template

[See `ExcelWriterApiTest.testWriteExcelTemplate()`](../src/test/java/com/savdev/commons/excel/api/ExcelWriterApiTest.java) 

```java
  try (
    InputStream stream = testPathUtils.testInputStream(TEMPLATE);
    FileOutputStream outputStream = new FileOutputStream(tempFile)) {
    excelWriterApi.writeExcelTemplate(
      stream,
      EXCEL_SHEET_NAME,
      HEADER_LINE_NUMBER + 1,
      excelLines(),
      outputStream);
  }
```

### Write into Excel without template

[See `ExcelWriterApiTest.testWriteExcel()`](../src/test/java/com/savdev/commons/excel/api/ExcelWriterApiTest.java)

```java
  try (FileOutputStream outputStream = new FileOutputStream(tempFile)) {
    excelWriterApi.writeExcel(
      EXCEL_SHEET_NAME,
      1, //no template, no header
      excelLines(),
      outputStream);
  }
```

### Writing without having Excel sheet name defined

Pass `null` as `sheetName` parameter. 

[See `ExcelWriterApiTest.testWriteExcelTemplateWithoutSheetName()`](../src/test/java/com/savdev/commons/excel/api/ExcelWriterApiTest.java)

```java
  try (
    InputStream stream = testPathUtils.testInputStream(TEMPLATE);
    FileOutputStream outputStream = new FileOutputStream(tempFile)) {
    excelWriterApi.writeExcelTemplate(
      stream,
      null, // do not set sheet name, the 1st sheet is used by default
      HEADER_LINE_NUMBER + 1,
      excelLines(),
      outputStream);
  }
```