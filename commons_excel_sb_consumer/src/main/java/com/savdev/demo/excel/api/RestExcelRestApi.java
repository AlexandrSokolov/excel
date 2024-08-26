package com.savdev.demo.excel.api;


import jakarta.ws.rs.Consumes;
import jakarta.ws.rs.POST;
import jakarta.ws.rs.Path;
import jakarta.ws.rs.Produces;
import jakarta.ws.rs.core.MediaType;

import jakarta.ws.rs.core.Response;
import org.jboss.resteasy.plugins.providers.multipart.MultipartFormDataInput;

import java.util.List;
import java.util.Map;

@Path(RestExcelRestApi.EXCEL_REST_PATH)
@Produces({MediaType.APPLICATION_JSON})
public interface RestExcelRestApi {

  String EXCEL_XLS_MEDIA_TYPE = "application/vnd.ms-excel";
  String EXCEL_XLSX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  String EXCEL_REST_PATH = "/excel";

  /**
   * Returns parsed Excel lines
   *
   * @param input
   * @return
   */
  @Path("/upload")
  @POST
  @Consumes(MediaType.MULTIPART_FORM_DATA)
  List<Map<String, Object>> uploadFile(MultipartFormDataInput input);

  @POST
  @Path("/download")
  @Produces(EXCEL_XLSX_MEDIA_TYPE)
  @Consumes(MediaType.APPLICATION_JSON)
  Response downloadFile(List<Map<String, Object>> lines);
}
