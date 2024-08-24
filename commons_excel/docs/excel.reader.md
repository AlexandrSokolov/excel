
The purpose of this Excel reader is - to allow users to work with Excel lines in stream api way

### Options to stream excel:

- [pre-configured mapping between Excel column letter(s) and named attribute](#pre-configured-excel-column-letters-to-named-attribute-mapping)
- [configuration extracted from Excel line](#configuration-extracted-from-excel-line)
- [with no-configuration, you access fields by column letters](#no-configuration-you-access-fields-by-column-letters)
- [covert stream of lines to Java POJO with transformers](#stream-of-excel-lines-as-java-pojos-with-transformer)
- [covert stream of lines to Java types manually](#stream-of-excel-lines-as-java-objects-manually)

### Pre-configured Excel column letter(s) to named attribute mapping

You want in your application access column value of each line via named attribute, like:
```java
var price = line.get("money");
```
[You have to](../src/test/java/com/savdev/commons/excel/config/TestExcelReaderConfig.java):
- configure the mapping between Excel columns (A, B, C, etc) to the named attributes as:
  ```java
  final static Map<String, String> configuration = Map.of(
    "A", "colA",
    "B", "colB",
    "E", "money",
    "F", "persentage",
    "G", "someDate",
    "H", "time",
    "I", "number");
  ```
- create Excel Reader with the configuration applied:
  ```java
    public ExcelReaderApi excelReaderApi() {
      return ExcelReaderService.instance(configuration);
    }
  ```
[Usage example, see `ExcelReaderApiTest.testTransformer()](../src/test/java/com/savdev/commons/excel/api/ExcelReaderApiTest.java):
- get a stream via `linesStream()`
- filter all lines, that do not belong to data to iterate, if exists
- use `ExcelReaderApi.transformer` mapper to get each Excel line as a mapping between the named attribute 
  and its value - in Excel column

```java
try (InputStream stream = testPathUtils.testInputStream(EXCEL_FILE)) {
  List<Map<String, Object>> result = excelReaderApi.linesStream(EXCEL_SHEET_NAME, stream)
    .filter(excelLine -> excelLine.getExcelLineNumber() > HEADER_LINE_NUMBER)
    .map(excelReaderApi::transformer)
    .toList();
  ...
}
```

### Configuration extracted from Excel line

You Excel contains line, that instead of values contains named attributes for better understanding:

| A            | B            | C            |
|:-------------|:-------------|:-------------|
| Attribute #1 | Attribute #2 | Attribute #3 |
| value A - 1  | value B - 1  | value C - 1  |
| value A - 2  | value B - 2  | value C - 2  |
| value A - 3  | value B - 3  | value C - 3  |

Instead of explicit configuration, like:
  ```java
  final static Map<String, String> configuration = Map.of(
    "A", "Attribute #1",
    "B", "Attribute #2",
    "C", "Attribute #3");
  ```

You extract the configuration from the 1st line.

Purpose: customer can move columns in the future. In your code you use only attribute names, to access value.

TODO

### No-configuration, you access fields by column letters

For certain cases, it might be enough to refer to columns via letters. 

[Example, see in `ExcelReaderApiTest.testNonConfiguredReader()`](../src/test/java/com/savdev/commons/excel/api/ExcelReaderApiTest.java):
```java
try (InputStream stream = testPathUtils.testInputStream(EXCEL_FILE)) {
  var lines = ExcelReaderService.instance()
    .linesStream(EXCEL_SHEET_NAME, stream)
    .peek(line -> logger.info("column B value: " +
      line.getColumnName2Attribute2Value()
        .get("B")
        .getValue()))
    .toList();
}
```

### Stream of Excel lines as Java POJOs with transformer

[Usage example, see `ExcelReaderApiTest.testTransformer2Dto()](../src/test/java/com/savdev/commons/excel/api/ExcelReaderApiTest.java):
- [create Java POJO, that contains fields, setters and getters, that refer to the named attributes/fields](../src/test/java/com/savdev/commons/excel/api/TestDto.java)
- get a stream via `linesStream()`
- filter all lines, that do not belong to data to iterate, if exists
- use `ExcelReaderApi.transformer` mapper to get each Excel line as a mapping between the named attribute
  and its value - in Excel column
- use `ExcelReaderApi.transformer` mapper that accepts the POJO supplier and the mapping, we get in the previous step

```java
try (InputStream stream = testPathUtils.testInputStream(EXCEL_FILE)) {
  List<Map<String, Object>> result = excelReaderApi.linesStream(EXCEL_SHEET_NAME, stream)
    .filter(excelLine -> excelLine.getExcelLineNumber() > HEADER_LINE_NUMBER)
    .map(excelReaderApi::transformer) 
    .map(attribute2Value -> excelReaderApi.transformer(
      TestDto::new,
      attribute2Value))
    .toList();
  ...
}
```

### Stream of Excel lines as Java objects manually

It might be an issue to create DTO for complex cases. 
For instance, you might want to use:
- Java record instead of POJO or
- Excel value might contain list of values you want to parse manually

In this case you are responsible for the mapping.




- [create configuration of `ExcelReaderService` and the service itself](#create-configuration-of-excelreaderservice-and-the-service-itself)
- [create unconfigurable `ExcelReaderService`](#create-unconfigurable-excelreaderservice)
- [`ExcelLine` class/abstraction](#excelline-classabstraction)
- [Handling excel lines via a stream](#handling-excel-lines-via-a-stream)
- [Filtering and validation](#filtering-and-validation)
- [Converting stream of `ExcelLine` into attribute -> value map](#converting-stream-of-excelline-into-attribute---value-map)
- [Converting stream of `ExcelLine` into stream of DTO objects](#converting-stream-of-excelline-into-stream-of-dto-objects)


#### Create configuration of `ExcelReaderService` and the service itself

The configuration it is just a mapping between column name in Excel (letters `A`, `B`, etc.)
and attribute name, which has some business meaning, for instance attribute in a REST DTO object.


#### Create unconfigurable `ExcelReaderService`

```java
public class ExcelReaderServiceTest {
  
  ExcelReaderService nonConfiguredService = ExcelReaderService.instance();
```

#### `ExcelLine` class/abstraction

All information about excel line is stored in `ExcelLine` class/abstraction`.
It contains:
1. line number according to its position in the Excel file.
2. An Excel column name -> attribute name -> value mapping.

The value could be any of the following Java Data Type:
- `BigDecimal`
- `String`
- `LocalDateTime` (also dates are defined as `LocalDateTime`)
- `LocalTime`
- `Boolean`

#### Handling excel lines via a stream

If all data is on the 1st tab, you might use unnamed version of the method,
Or use a name of an Excel sheet to extract data from it:

```java
public class ExcelReaderServiceTest {
  
  ExcelReaderService service = ...

  @Test
  public void testReadExcelWithSheetName() throws IOException {
    try (InputStream stream = testInputStream(EXCEL_FILE)) {
      Stream<ExcelLine> excelStream = service.linesStream(stream);
      Stream<ExcelLine> excelStreamWithNamedSheet = service.linesStream(EXCEL_SHEET_NAME, stream);
      Assertions.assertNotNull(excelStream);
      Assertions.assertEquals(4, excelStream.count());
  }

```

#### Filtering and validation

In the following example, we skip handling header (the first 3 lines),
Then we split lines into valid and invalid, to handle them differently.

Invalid lines are those, which have empty value in `D` column:

```java
  @Test
  public void testFilteringAndValidation() throws IOException {
    try (InputStream stream = testInputStream(EXCEL_FILE)) {

      Map<Boolean, List<ExcelLine>> validAndInvalid = configuredService.linesStream(stream)
        .filter(excelLine -> excelLine.getExcelLineNumber() > 3)
        .collect(Collectors.partitioningBy(excelLine -> excelLine.valueByColumn("D")
          .map(value -> (String) value)
          .map(StringUtils::isNoneEmpty)
          .isPresent()));
      Assertions.assertEquals(2, validAndInvalid.size());
      Assertions.assertEquals(1, validAndInvalid.get(true).size());
      Assertions.assertEquals(1, validAndInvalid.get(false).size());
    }
  }
```

#### Converting stream of `ExcelLine` into attribute -> value map

```java
  @Test
  public void testTransformer() throws IOException {
    try (InputStream stream = testInputStream(EXCEL_FILE)) {
      List<Map<String, Object>> result = configuredService.linesStream(stream)
        .filter(excelLine -> excelLine.getExcelLineNumber() > 3)
        .map(configuredService::transformer)
        .collect(Collectors.toList());
      Assertions.assertEquals(2, result.size());
    }
  }
```

#### Converting stream of `ExcelLine` into stream of DTO objects

In the following example we:
1. skip the header (the first 3 lines)
2. transform `ExcelLine` into map
3. transform map into `TestDto`

```java
  @Test
  public void testTransformer2Dto() throws IOException {
    try (InputStream stream = testInputStream(EXCEL_FILE)) {
      List<TestDto> result = configuredService.linesStream(EXCEL_SHEET_NAME, stream)
        .filter(excelLine -> excelLine.getExcelLineNumber() > 3)
        .map(configuredService::transformer)
        .map(attribute2Value -> configuredService.transformer(
            TestDto::new,
            attribute2Value))
        .collect(Collectors.toList());
      Assertions.assertEquals(2, result.size());
    }
  }
```

`TestDto` example:
```java
public class TestDto {
  String colA;
  BigDecimal money;
  LocalDateTime someDate;
  LocalTime time;
  BigDecimal number;
  // getters and setters
}
```