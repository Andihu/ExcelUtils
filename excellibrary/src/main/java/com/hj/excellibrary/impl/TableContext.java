package com.hj.excellibrary.impl;


import com.hj.excellibrary.service.ITabContext;

import java.util.Map;

public class TableContext implements ITabContext {

    private String sheetName;

    private int sheetIndex;

    private Map<Integer, String> excelProperty;

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public void setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public void setExcelProperty(Map<Integer, String> excelProperty) {
        this.excelProperty = excelProperty;
    }

    @Override
    public String getSheetName() {
        return sheetName;
    }

    @Override
    public int getSheetIndex() {
        return sheetIndex;
    }


    @Override
    public Map<Integer, String> getExcelProperty() {
        return excelProperty;
    }

    @Override
    public String toString() {
        return "TableContext{" +
                "sheetName='" + sheetName + '\'' +
                ", sheetIndex=" + sheetIndex +
                ", excelProperty=" + excelProperty +
                '}';
    }
}
