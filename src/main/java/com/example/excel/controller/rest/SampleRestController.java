package com.example.excel.controller.rest;

import com.example.excel.service.SampleService;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

@RestController
@RequiredArgsConstructor
@Slf4j
public class SampleRestController {
  private final SampleService sampleService;

  @GetMapping("/download")
  public void download(HttpServletResponse res) throws Exception {
    sampleService.write(res);
  }

  @PostMapping("/upload")
  public void upload(MultipartHttpServletRequest request) throws Exception {
    log.info("test {}", request.getParameter("test"));

    MultipartFile excelFile = request.getFile("file");
    String password = request.getParameter("password");
    sampleService.read(excelFile.getInputStream(), password);

  }
}
