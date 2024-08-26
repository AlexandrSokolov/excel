### Topics:

- [Adding Commons Excel library into the application](#adding-commons-excel-library-into-the-application)
- [Run and test the application](#run-and-test-the-application)

### Adding Commons Excel library into the application

1. Add dependency on the library
  ```xml
      <dependency>
        <groupId>com.savdev.commons.excel</groupId>
        <artifactId>commons-excel</artifactId>
        <version>${commons-excel.version}</version>
      </dependency>
  ```
2. [Add Excel Reader configuration, if you need file reading](src/main/java/com/savdev/demo/excel/config/ExcelServicesConfiguration.java)
3. [Add Excel Service that exposes to the application the simplified API](src/main/java/com/savdev/demo/excel/service/DemoExcelService.java)
4. [TODO add integration tests for excel service]

### Adding rest endpoints if needed

Add Rest API and its implementation for Excel uploading/downloading, if you provide such functionality.

1. [Rest Api](src/main/java/com/savdev/demo/excel/api/RestExcelRestApi.java)
2. [Rest Service](src/main/java/com/savdev/demo/excel/rest/service/ExcelRestService.java)
3. [TODO add integration tests for file uploading/downloading]

### Run and test the application

Run:
```bash
mvn spring-boot:run
```

Test file downloading, see result in `/target/t.xlsx` file:
```bash
cd commons_excel_sb_consumer
curl -J -X POST -w "\n" -H 'Content-Type: application/json' -d @./src/test/resources/lines.json -o ./target/t.xlsx http://localhost:8080/rest/excel/download
```

Test file uploading and getting back the extracted lines:
```bash
curl -w "\n" -F "fileName=test.xlsx" -F file=@./target/t.xlsx http://localhost:8080/rest/excel/upload
```