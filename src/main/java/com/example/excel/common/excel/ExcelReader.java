package com.example.excel.common.excel;

import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.util.Strings;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelReader {

  /**
   * rowData 데이터가 모두 빈 문자열이 아닌지 체크한다.
   *
   * @param rowData
   */
  private boolean validationRowData(Map<String, String> rowData) {
    return rowData.values().stream().anyMatch(Strings::isNotEmpty);
  }

  /**
   * Cell에 해당하는 Column Name을 가져온다(A,B,C..)
   *
   * @param cell
   * @param index
   */
  private String getCellName(Cell cell, int index) {
    int cellNum = 0;
    if (cell != null) {
      cellNum = cell.getColumnIndex();
    } else {
      cellNum = index;
    }

    return CellReference.convertNumToColString(cellNum);
  }

  /**
   * cell에 데이터를 하나에 row 데이터로 리턴한다.
   *
   * @param row
   * @param option
   */
  private Map<String, String> getRowData(XSSFRow row, ReadOption option) {
    Map<String, String> rowData = new HashMap<>();
    DataFormatter formatter = new DataFormatter();

    for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
      XSSFCell cell = row.getCell(cellIndex);

      if (cell == null) {
        continue;
      }

      String key = this.getCellName(cell, cellIndex);
      String value = formatter.formatCellValue(cell);

      if (option.contains(key)) {
        rowData.put(key, value);
      }
    }

    return rowData;
  }

  /**
   * 첫 번째 시트만 읽어드린다.
   *
   * @param option 옵션
   * @param is     (InputStream)
   * @throws IOException
   */
  public List<Map<String, String>> read(ReadOption option, InputStream is) throws Exception {
    List<Map<String, String>> rows = new ArrayList<>();

    XSSFWorkbook workbook = null;

    if (StringUtils.isNotEmpty(option.getPassword())) {
      POIFSFileSystem fs = new POIFSFileSystem(is);
      EncryptionInfo info = new EncryptionInfo(fs);
      Decryptor decryptor = Decryptor.getInstance(info);

      if (decryptor.verifyPassword(option.getPassword())) {
        workbook = new XSSFWorkbook(decryptor.getDataStream(fs));
      } else {
        throw new Exception("wrong the password");
      }
    } else {
      workbook = new XSSFWorkbook(is);
    }

    XSSFSheet sheet = workbook.getSheetAt(0);
    is.close();

    int rowIndex = option.getRowIndex();

    for (; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
      XSSFRow row = sheet.getRow(rowIndex);
      Map<String, String> rowData = this.getRowData(row, option);

      if (this.validationRowData(rowData)) {
        rows.add(rowData);
      }
    }

    return rows;
  }

  /**
   * 로우에서 추출할 셀에 키 값이 모두 존재하는지 여부
   *
   * @param row
   * @param option
   */
  public boolean containsKeys(Map<String, String> row, ReadOption option) {
    return row.keySet().containsAll(option.getCellNames());
  }
}
