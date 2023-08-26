package com.example.excel.common.excel;

import org.apache.commons.lang3.ObjectUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.List;

public abstract class Excel<T> {
  private final String SHEET_NAME = "Sheet";

  // 시트에 최대한 쓸 수 있는 row 수
  private final int MAX_ROW = 1040000;
  private List<String> header = null;
  private List<T> data = null;

  public Excel(List<String> header, List<T> data) {
    this.header = header;
    this.data = data;
  }

  public void write(SXSSFWorkbook workbook, int rowNum) {
    // 매개변수로 받은 rowNo부터 이어서 작성
    int rowNo = rowNum % MAX_ROW;
    int rowDataNo = rowNo;

    if (ObjectUtils.isNotEmpty(data)) {
      for (T t : data) {
        String sheetName = SHEET_NAME + (rowNo / MAX_ROW + 1);
        boolean isNeedHeader = this.isNeedHeader(workbook, sheetName);
        SXSSFSheet sheet = this.getSheet(workbook, sheetName);

        if (isNeedHeader) {
          drawHeader(workbook, sheet, header);
          rowDataNo = (rowNo + 1) % MAX_ROW;
        }

        drawData(workbook, sheet, rowDataNo, t);
        rowNo = rowNo + 1;
        rowDataNo = rowDataNo + 1;
      }
    } else {
      // 데이터가 없을 경우 헤더라도 그릴 수 있도록 처리.
      String sheetName = SHEET_NAME + (rowNo / MAX_ROW + 1);

      boolean isNeedHeader = this.isNeedHeader(workbook, sheetName);
      SXSSFSheet sheet = this.getSheet(workbook, sheetName);

      if (isNeedHeader) {
        drawHeader(workbook, sheet, header);
        rowNo = rowNo + 1;
      }
    }
  }

  protected abstract void drawHeader(SXSSFWorkbook workbook, SXSSFSheet sheet, List<String> header);

  protected abstract void drawData(SXSSFWorkbook workbook, SXSSFSheet sheet, int rowNo, T data);

  // 헤더 셀 스타일
  protected CellStyle makeHeadStyle(SXSSFWorkbook workbook) {
    CellStyle style = makeBorder(workbook);
    style.setFillForegroundColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    style.setAlignment(HorizontalAlignment.CENTER);
    style.setVerticalAlignment(VerticalAlignment.CENTER);

    Font font = workbook.createFont();
    font.setBold(true);

    style.setWrapText(true);
    style.setFont(font);
    return style;
  }


  private CellStyle makeBorder(SXSSFWorkbook workbook) {
    CellStyle style = workbook.createCellStyle();
    style.setBorderTop(BorderStyle.THIN);
    style.setBorderBottom(BorderStyle.THIN);
    style.setBorderLeft(BorderStyle.THIN);
    style.setBorderRight(BorderStyle.THIN);
    style.setVerticalAlignment(VerticalAlignment.CENTER);
    style.setAlignment(HorizontalAlignment.CENTER);

    return style;
  }

  // 셀 스타일 밑줄
  protected CellStyle makeBottomBorder(SXSSFWorkbook workbook) {
    CellStyle style = workbook.createCellStyle();
    style.setBorderBottom(BorderStyle.THIN);

    return style;
  }

  // 셀 스타일 왼쪽 줄
  protected CellStyle makeLeftBorder(SXSSFWorkbook workbook) {
    CellStyle style = workbook.createCellStyle();
    style.setBorderLeft(BorderStyle.THIN);

    return style;
  }

  // 셀 스타일 오른쪽 줄
  protected CellStyle makeRightBorder(SXSSFWorkbook workbook) {
    CellStyle style = workbook.createCellStyle();
    style.setBorderRight(BorderStyle.THIN);

    return style;
  }

  // 셀 스타일 윗줄
  protected CellStyle makeTopBorder(SXSSFWorkbook workbook) {
    CellStyle style = workbook.createCellStyle();
    style.setBorderTop(BorderStyle.THIN);

    return style;
  }

  // 셀 스타일 왼쪽, 밑줄
  protected CellStyle makeLeftBottomBorder(SXSSFWorkbook workbook) {
    CellStyle style = workbook.createCellStyle();
    style.setBorderLeft(BorderStyle.THIN);
    style.setBorderBottom(BorderStyle.THIN);

    return style;
  }

  // 셀 스타일 오른쪽, 밑줄
  protected CellStyle makeRightBottomBorder(SXSSFWorkbook workbook) {
    CellStyle style = workbook.createCellStyle();
    style.setBorderRight(BorderStyle.THIN);
    style.setBorderBottom(BorderStyle.THIN);

    return style;
  }


  private SXSSFSheet getSheet(SXSSFWorkbook workbook, String sheetName) {
    SXSSFSheet sheet = workbook.getSheet(sheetName);

    if (ObjectUtils.isEmpty(sheet)) {
      sheet = workbook.createSheet(sheetName);
    }

    sheet.trackAllColumnsForAutoSizing();

    return sheet;
  }

  private boolean isNeedHeader(SXSSFWorkbook workbook, String sheetName) {
    SXSSFSheet sheet = workbook.getSheet(sheetName);
    return ObjectUtils.isEmpty(sheet);
  }
}
