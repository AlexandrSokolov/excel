package com.savdev.demo.excel.rest.service;

import com.savdev.demo.excel.api.RestExcelRestApi;
import com.savdev.demo.excel.service.DemoExcelService;
import jakarta.ws.rs.core.Response;
import jakarta.ws.rs.core.StreamingOutput;
import org.jboss.resteasy.plugins.providers.multipart.MultipartFormDataInput;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

@Service
public class ExcelRestService implements RestExcelRestApi {

  @Autowired
  private DemoExcelService demoExcelService;

  @Override
  public List<Map<String, Object>> uploadFile(MultipartFormDataInput input) {
    var fileUploadDto = FileUploadDto.instance(input);
    try (InputStream inputStream = fileUploadDto.inputStream()) {
      return demoExcelService.extractLines(inputStream);
    } catch (IOException e) {
      throw new IllegalStateException(e);
    }
  }

  @Override
  public Response downloadFile(List<Map<String, Object>> lines) {
    return Response.ok()
      .entity((StreamingOutput) outputStream -> {
        demoExcelService.downloadAsFile(lines, outputStream);
        outputStream.flush();
      })
      .build();
  }
}
