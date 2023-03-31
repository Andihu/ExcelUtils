package com.hj.excellibrary.impl;

import android.os.Handler;
import android.os.Looper;
import android.util.Log;


import com.hj.excellibrary.NotFindSheetException;
import com.hj.excellibrary.annotation.CellStyle;
import com.hj.excellibrary.annotation.ExcelTable;
import com.hj.excellibrary.annotation.ExcelWriteAdapter;
import com.hj.excellibrary.annotation.ExcelWriteCell;
import com.hj.excellibrary.service.IConvertParserAdapter;
import com.hj.excellibrary.service.IExcelUtils;
import com.hj.excellibrary.service.IParseListener;
import com.hj.excellibrary.service.IParser;
import com.hj.excellibrary.service.ITabContext;
import com.hj.excellibrary.service.IWriteListener;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.List;

public class ExcelImpl implements IExcelUtils {

    private static final String TAG = "ExcelImpl";

    private final Handler handler = new Handler(Looper.getMainLooper());

    @Override
    public <T> void readExcelXSSF(InputStream inputStream, Class<T> aClass, IParseListener<T> listener) {
        IParser<T> parser = new ParserImpl<>();
        new Thread(() -> {
            try {
                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
                doReadExcelXLSX(parser, workbook, aClass, listener);
            } catch (IOException|NotFindSheetException e) {
                e.printStackTrace();
                handler.post(() -> listener.onParseError(e));
            }finally {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }).start();

    }

    @Override
    public <T> void readExcelHSSF(InputStream inputStream, Class<T> tClass, IParseListener<T> listener) {
        IParser<T> parser = new ParserImpl<>();
        new Thread(() -> {
            try {
                HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
                doReadExcelXLSX(parser, workbook, tClass, listener);
            } catch (IOException|NotFindSheetException e) {
                e.printStackTrace();
                handler.post(() -> listener.onParseError(e));
            }finally {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }).start();
    }

    @Override
    public <T> void writeExcelXSSF(OutputStream outputStream, List<T> data, IWriteListener listener) {
        if (data == null || data.size() <= 0) {
            return;
        }
        // 写入表头
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet;
        try {
            sheet = writeTableHead(workbook, data.get(0));
        } catch (IllegalAccessException | InstantiationException e) {
            e.printStackTrace();
            handler.post(() -> listener.onWriteError(e));
            return;
        }
        //写入数据
        for (int i = 0; i < data.size(); i++) {
            XSSFRow row = sheet.createRow(i + 1);
            T t = data.get(i);
            for (Field declaredField : t.getClass().getDeclaredFields()) {
                ExcelWriteCell excelWriteCell = declaredField.getAnnotation(ExcelWriteCell.class);
                CellStyle cellStyle = declaredField.getAnnotation(CellStyle.class);
                if (excelWriteCell != null) {
                    XSSFCell cell = row.createCell(excelWriteCell.writeIndex());
                    try {
                        cell.setCellValue(String.valueOf(declaredField.get(t)));
                        setCellStyle(workbook, cell, cellStyle);
                    } catch (IllegalAccessException e) {
                        handler.post(() -> listener.onWriteError(e));
                    }

                }
            }
            for (Field declaredField : t.getClass().getDeclaredFields()) {
                ExcelWriteAdapter excelWriteAdapter = declaredField.getAnnotation(ExcelWriteAdapter.class);
                CellStyle cellStyle = declaredField.getAnnotation(CellStyle.class);
                if (excelWriteAdapter != null) {
                    try {
                        IConvertParserAdapter adapter = (IConvertParserAdapter) excelWriteAdapter.adapter().newInstance();
                        adapter.convert((column, value, columnIndex) -> {
                            XSSFCell cell = row.createCell(columnIndex);
                            cell.setCellValue(String.valueOf(value));
                            setCellStyle(workbook, cell, cellStyle);
                        }, declaredField.get(t));
                    } catch (IllegalAccessException | InstantiationException e) {
                        handler.post(() -> listener.onWriteError(e));
                    }
                }
            }
        }


        try {
            sheet.setColumnWidth(0,10000);
            workbook.write(outputStream);
            outputStream.flush();
            if (listener != null) {
                handler.post(listener::onEndWrite);
            }
        } catch (Exception e) {
            e.printStackTrace();
            if (listener != null) {
                handler.post(() -> listener.onWriteError(e));
            }
        } finally {
            try {
                if (outputStream != null)
                    outputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }


    }

    private void setCellStyle(XSSFWorkbook workbook, XSSFCell cell, CellStyle cellStyle) {
        if (cellStyle == null) {
            return;
        }
        boolean bold = cellStyle.bold();
        boolean italic = cellStyle.italic();
        FontUnderline underline = cellStyle.underline();
        HorizontalAlignment horizontalAlignment = cellStyle.horizontalAlign();
        VerticalAlignment verticalAlignment = cellStyle.verticalAlign();
        XSSFCellStyle cellStyle1 = workbook.createCellStyle();
        cellStyle1.setVerticalAlignment(verticalAlignment);
        cellStyle1.setAlignment(horizontalAlignment);
        XSSFFont font = workbook.createFont();
        font.setBold(bold);
        font.setItalic(italic);
        font.setUnderline(underline);
        cellStyle1.setFont(font);
        cell.setCellStyle(cellStyle1);
    }

    private <T> XSSFSheet writeTableHead(XSSFWorkbook workbook, T t) throws IllegalAccessException, InstantiationException {
        XSSFSheet sheet;
        ExcelTable excelTableAnnotation = t.getClass().getAnnotation(ExcelTable.class);
        if (excelTableAnnotation != null) {
            String s = excelTableAnnotation.sheetName();
            if (s != null && s.length() > 0) {
                sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName(s));
            } else {
                sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName("Sheet1"));
            }
        } else {
            sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName("Sheet1"));
        }
        XSSFRow row = sheet.createRow(0);
        Field[] declaredFields = t.getClass().getDeclaredFields();
        for (Field declaredField : declaredFields) {
            ExcelWriteCell annotation = declaredField.getAnnotation(ExcelWriteCell.class);
            CellStyle cellStyle = declaredField.getAnnotation(CellStyle.class);
            if (annotation != null) {
                Log.d(TAG, "writeTableHead: writeName = " + annotation.writeName());
                Log.d(TAG, "writeTableHead: writeIndex = " + annotation.writeIndex());
                XSSFCell cell = row.createCell(annotation.writeIndex());
                cell.setCellValue(annotation.writeName());
                setCellStyle(workbook, cell, cellStyle);
            }
        }
        for (Field declaredField : declaredFields) {
            ExcelWriteAdapter excelWriteAdapter = declaredField.getAnnotation(ExcelWriteAdapter.class);
            CellStyle cellStyle = declaredField.getAnnotation(CellStyle.class);
            if (excelWriteAdapter != null) {
                IConvertParserAdapter adapter = (IConvertParserAdapter) excelWriteAdapter.adapter().newInstance();
                adapter.convert((column, value, columnIndex) -> {
                    Log.d(TAG, "writeTableHead: writeName = " + column);
                    Log.d(TAG, "writeTableHead: writeIndex = " + columnIndex);
                    XSSFCell cell = row.createCell(columnIndex);
                    cell.setCellValue(String.valueOf(column));
                    setCellStyle(workbook, cell, cellStyle);
                }, declaredField.get(t));
            }
        }
        return sheet;
    }

    private <T> void doReadExcelXLSX(IParser<T> parser, Workbook workbook, Class<T> aClass, IParseListener<T> listener) throws NotFindSheetException {
        ITabContext tabContexts = parser.parseXSSFContext(workbook, aClass);
        Log.d(TAG, "doReadExcelXLSX: tabContexts = " + tabContexts.toString());
        parser.doParse(workbook, tabContexts, aClass, listener);
    }
}
