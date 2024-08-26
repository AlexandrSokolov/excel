package com.savdev.demo.excel.rest.service;

import org.jboss.resteasy.plugins.providers.multipart.InputPart;
import org.jboss.resteasy.plugins.providers.multipart.MultipartFormDataInput;

import jakarta.ws.rs.core.HttpHeaders;
import jakarta.ws.rs.core.MultivaluedMap;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

public record FileUploadDto (
  String fileName,
  //String contentType, //todo
  InputStream inputStream
) {
  public static FileUploadDto instance(MultipartFormDataInput input) {
    try {
      Map<String, List<InputPart>> uploadForm = input.getFormDataMap();

      if (!uploadForm.containsKey("file")){
        throw new IllegalStateException("'file' attribute is expected");
      }

      List<InputPart> inputParts = uploadForm.get("file");

      if (inputParts.size() != 1) {
        throw new IllegalStateException("Only a single part is expected");
      }

      InputPart inputPart = inputParts.getFirst();

      MultivaluedMap<String, String> header = inputPart.getHeaders();

      InputStream inputStream = inputPart.getBody(InputStream.class,null);

      return new FileUploadDto(
        extractFileName(header),
        inputStream);

    } catch (IOException e) {
      throw new IllegalStateException("Could not extract file data", e);
    }
  }

  /**
   * header sample
   * {
   * 	Content-Type=[image/png],
   * 	Content-Disposition=[form-data; name="file"; filename="filename.extension"]
   * }
   **/
  //get uploaded filename, is there a easy way in RESTEasy?
  private static String extractFileName(MultivaluedMap<String, String> header) {

    String[] contentDisposition = header.getFirst(HttpHeaders.CONTENT_DISPOSITION).split(";");

    for (String filename : contentDisposition) {
      if (filename.trim().startsWith("filename")) {

        String[] name = filename.split("=");

        return name[1].trim().replaceAll("\"", "");
      }
    }
    return "unknown";
  }
}
