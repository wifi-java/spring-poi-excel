package com.example.excel.service.excel;

import com.example.excel.common.excel.Excel;
import com.example.excel.model.User;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.List;

public class ExcelUser extends Excel<User> {

  public ExcelUser(List<String> header, List<User> data) {
    super(header, data);
  }

  @Override
  protected void drawHeader(SXSSFWorkbook workbook, SXSSFSheet sheet, List<String> header) {
    if (ObjectUtils.isNotEmpty(header)) {
      CellStyle headStyle = makeHeadStyle(workbook);
      Row row = sheet.createRow(0);

      for (int i = 0; i < header.size(); i++) {
        String title = header.get(i);

        Cell cell = row.createCell(i);
        cell.setCellStyle(headStyle);
        cell.setCellValue(title);
      }
    }
  }

  @Override
  protected void drawData(SXSSFWorkbook workbook, SXSSFSheet sheet, int rowNo, User data) {
    Row row = sheet.createRow(rowNo);

    Cell nameCell = row.createCell(0, CellType.STRING);
    nameCell.setCellValue(data.getName());

    Cell emailCell = row.createCell(1, CellType.STRING);
    emailCell.setCellValue(data.getEmail());

    CellStyle cellLeftStyle = makeLeftBottomBorder(workbook);
    CellStyle cellRightStyle = makeRightBottomBorder(workbook);
    nameCell.setCellStyle(cellLeftStyle);
    emailCell.setCellStyle(cellRightStyle);

    // 셀 자동 사이즈 조절
    sheet.autoSizeColumn(0);
    sheet.autoSizeColumn(1);

    // 셀 여백 조절
    sheet.setColumnWidth(0, (sheet.getColumnWidth(0)) + (short) 512);
    sheet.setColumnWidth(1, (sheet.getColumnWidth(1)) + (short) 512);
  }
}
