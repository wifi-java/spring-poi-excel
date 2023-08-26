package com.example.excel.common.excel;

import java.util.ArrayList;
import java.util.List;

public class ReadOption {
    /**
     * 엑셀 비밀번호
     */
    private String password;

    /**
     * 몇 행부터 읽어드릴지에 대한 index
     */
    private int rowIndex;

    /**
     * 추출한 cell 설정
     */
    private List<String> cellNames = new ArrayList<>();

    public ReadOption(int rowIndex, String... cellNames) {
        this.setRowIndex(rowIndex);
        this.setOutputCell(cellNames);
    }

    /**
     * index가 0보다 작은 값이 들어오면 0으로 설정해주고 그렇지 않으면 전달받은 값을 리턴해준다.
     *
     * @param index
     */
    private int getValidationIndex(int index) {
        return Math.max(index, 0);
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = this.getValidationIndex(rowIndex);
    }

    public void setOutputCell(String... cellNames) {
        this.cellNames.clear();

        if (cellNames != null) {
            for (String cellName : cellNames) {
                if (!this.cellNames.contains(cellName)) {
                    this.cellNames.add(cellName);
                }
            }
        }
    }

    public String getPassword() {
        return password;
    }

    public void setPassword(String password) {
        this.password = password;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public List<String> getCellNames() {
        return cellNames;
    }

    public boolean contains(String cellName) {
        return this.cellNames.contains(cellName);
    }

}
