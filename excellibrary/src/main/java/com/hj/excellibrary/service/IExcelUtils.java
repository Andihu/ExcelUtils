package com.hj.excellibrary.service;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

public interface IExcelUtils {

    /**
     * 读取Excel文件
     * 提供读写Microsoft Excel OOXML格式档案的功能
     * XSSF用于（*.xlsx）Excel文件
     * @param inputStream 文件流
     * @param tClass      解析的实体类
     * @param <T>         实体类范型
     * @throws IOException IO异常
     */
    <T> void readExcelXSSF(InputStream inputStream, Class<T> tClass,IParseListener<T> listener);

    /**
     * 读取Excel文件
     * 提供读写Microsoft Excel格式档案的功能
     * HSSF用于解析旧版本（*.xls）Excel文件，由于旧版本的Excel文件只能存在65535行数据，目前已经不常用
     * @param inputStream 文件流
     * @param tClass      解析的实体类
     * @param <T>         实体类范型
     * @throws IOException IO异常
     */
    <T> void readExcelHSSF(InputStream inputStream, Class<T> tClass,IParseListener<T> listener);

    /**
     * 输出Excel文件
     * 提供读写Microsoft Excel格式档案的功能
     * XSSF（*.xlsx）Excel文件
     * @param outputStream 文件输出流
     */
    <T> void writeExcelXSSF(OutputStream outputStream , List<T> data, IWriteListener listener);


}
