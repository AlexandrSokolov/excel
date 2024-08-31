### excel validator examples, value access by attribute name
```java
public class ValidFileNameValidator implements Validator<ExcelLine> {
@Override
public Optional<String> errorIfNotValid(ExcelLine excelLine) {

    var maybeValue = excelLine.getColumnName2Attribute2Value().values().stream()
                       .filter(entry -> FILE_NAME_COLUMN_NAME.equals(entry.getKey()))
                       .map(entry -> entry.getValue())
                       .findFirst();
```

### excel validator examples, converting to map

Note: we convert actually twice in validation, and later when stream the lines, so should be avoided in general:

```java
public class RequiredButEmptyColumnsValidator implements Validator<ExcelLine> {

  private static final ExcelColumnNames excelColumnNames = new ExcelColumnNames(){};

  @Autowired
  private ExcelReaderApi excelReaderApi;

  @Override
  public Optional<String> errorIfNotValid(ExcelLine excelLine) {
    var asMap = excelReaderApi.transformer(excelLine);
    var requiredButEmpty = excelColumnNames.requiredColumns().stream()
        .filter(column -> Objects.isNull(asMap.get(column)))
        .toList();
```

### Add example of using excel, not as a library, to simplify copy-paste for new projects