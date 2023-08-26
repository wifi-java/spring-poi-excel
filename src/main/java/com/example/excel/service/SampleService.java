package com.example.excel.service;

import com.example.excel.common.excel.Excel;
import com.example.excel.common.excel.ExcelReader;
import com.example.excel.common.excel.ExcelWriter;
import com.example.excel.common.excel.ReadOption;
import com.example.excel.model.User;
import com.example.excel.service.excel.ExcelUser;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.util.CollectionUtils;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Service
@Slf4j
public class SampleService {

  public void write(HttpServletResponse res) {
    try {
      List<String> header = new ArrayList<>();
      List<User> data = new ArrayList<>();

      header.add("이름");
      header.add("이메일");

      for (int i = 0; i < 10; i++) {
        User user = new User();
        user.setName(String.valueOf(i));
        user.setEmail(i + "@gmail.com");

        data.add(user);
      }

      // 샘플 엑셀 파일을 작성 한다.
      Excel<User> test = new ExcelUser(header, data);
      ExcelWriter excelWriter = new ExcelWriter(test);
      SXSSFWorkbook workbook = excelWriter.writeWorkbook();

      String fileName = "spring_excel_download";
      excelWriter.writeFile(workbook, fileName, res);

      // 엑셀 파일 암호화
      //String password = "1234";
      //excelWriter.writeFile(workbook, fileName, password, res);
    } catch (Exception e) {
      log.error("", e);
    }
  }

  public void read(InputStream inputStream, String password) throws Exception {
    ReadOption option = new ReadOption(2, "A", "B");
    if (StringUtils.isNotEmpty(password)) {
      option.setPassword(password);
    }

    ExcelReader reader = new ExcelReader();
    List<Map<String, String>> rows = reader.read(option, inputStream);

    if (!CollectionUtils.isEmpty(rows)) {
      for (Map<String, String> row : rows) {
        String name = row.get("A");
        String email = row.get("B");

        User user = new User();
        user.setName(name);
        user.setEmail(email);

        log.info("name: {}, email: {}", user.getName(), user.getEmail());
      }
    }
  }
}
