#### The purpose of this Excel Writer is - to hide excel-library specific logic and expose only standard java api types

**Note:** this Excel Writer uses the `column letter -> value` mapping, but not the `attribute -> value`

### How to use Excel Writer:

- [Write into Excel template](#write-into-excel-template)
- [Write into Excel without template](#write-into-excel-without-template)
- [Writing without having Excel sheet name defined](#writing-without-having-excel-sheet-name-defined)
- [Create Excel as InputStream to pass it further](#create-excel-as-inputstream-to-pass-it-further)

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

### Create Excel as InputStream to pass it further

[See `ExcelWriterApiTest.testWriteExcelIntoInputStream()`](../src/test/java/com/savdev/commons/excel/api/ExcelWriterApiTest.java)

```java
    try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
      excelWriterApi.writeExcel(
        EXCEL_SHEET_NAME,
        1, //no template, no header
        excelLines(),
        outputStream);
      try (InputStream inputStream = new ByteArrayInputStream(outputStream.toByteArray())) {
        var targetFile = tempFile();
        Files.copy(
          inputStream,
          targetFile.toPath(),
          StandardCopyOption.REPLACE_EXISTING);
        System.out.println(targetFile.getAbsolutePath());
      }
    }
```