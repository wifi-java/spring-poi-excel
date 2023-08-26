package com.example.excel.common.excel;

import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

@Slf4j
public class ExcelWriter {
  private Excel excel = null;

  public ExcelWriter(Excel excel) {
    this.excel = excel;
  }

  private SXSSFWorkbook getWorkbook(SXSSFWorkbook sxssfWorkbook) {
    // sxssfWorkbook 값이 존재하면 이어서 작성하고 없으면 새로운 객체를 생성
    if (ObjectUtils.isNotEmpty(sxssfWorkbook)) {
      return sxssfWorkbook;
    }

    return new SXSSFWorkbook(-1);
  }


  public SXSSFWorkbook writeWorkbook() {
    return this.writeWorkbook(null, 0);
  }

  public SXSSFWorkbook writeWorkbook(SXSSFWorkbook sxssfWorkbook, int rowNum) {
    SXSSFWorkbook workbook = getWorkbook(sxssfWorkbook);

    if (excel != null) {
      excel.write(workbook, rowNum);
    }

    return workbook;
  }

  // 암호화 없는 엑셀 파일 생성
  public void writeFile(SXSSFWorkbook workbook, String fileName, HttpServletResponse res) throws Exception {
    res.setHeader("Content-Disposition", "attachment;filename=" + fileName + ".xlsx");
    res.setContentType("application/msexcel;charset=UTF-8");

    try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
      workbook.write(baos);
      byte[] bytes = baos.toByteArray();
      workbook.close();

      res.setContentLength(bytes.length);
      res.getOutputStream().write(bytes);
    }
  }

  // 전달받은 비밀번호로 엑셀 파일 작성
  public void writeFile(SXSSFWorkbook workbook, String fileName, String password, HttpServletResponse res) throws Exception {
    res.setHeader("Content-Disposition", "attachment;filename=" + fileName + ".xlsx");
    res.setContentType("application/msexcel;charset=UTF-8");
    byte[] bytes;

    try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
      workbook.write(baos);
      bytes = baos.toByteArray();
      workbook.close();
    }

    try (InputStream is = new ByteArrayInputStream(bytes)) {
      try (POIFSFileSystem fs = new POIFSFileSystem()) {
        EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
        Encryptor encryptor = info.getEncryptor();

        encryptor.confirmPassword(password);

        try (OPCPackage opc = OPCPackage.open(is); OutputStream os = encryptor.getDataStream(fs)) {
          opc.save(os);
        }

        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
          fs.writeFilesystem(baos);
          bytes = baos.toByteArray();
          res.setContentLength(bytes.length);
          res.getOutputStream().write(bytes);
        }
      }
    }
  }
}
