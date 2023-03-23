package com.hj.excellibrary.service;


import com.hj.excellibrary.NotFindSheetException;

import org.apache.poi.ss.usermodel.Workbook;

public interface IParser<T> {

    /**
     * 解析数据
     */
    void doParse(Workbook workbook, ITabContext context, Class<T> tClass, IParseListener<T> listener);

    /**
     * 解析表头信息
     */
    ITabContext parseXSSFContext(Workbook workbook,Class<T> tClass) throws NotFindSheetException;


}
