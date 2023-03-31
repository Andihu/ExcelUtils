package com.hj.excellibrary.impl;

import android.annotation.SuppressLint;
import android.os.Handler;
import android.os.Looper;
import android.util.Log;

import com.hj.excellibrary.NotFindSheetException;
import com.hj.excellibrary.annotation.ExcelReadAggregate;
import com.hj.excellibrary.annotation.ExcelReadCell;
import com.hj.excellibrary.annotation.ExcelTable;
import com.hj.excellibrary.service.IParseListener;
import com.hj.excellibrary.service.IParser;
import com.hj.excellibrary.service.ITabContext;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Map;

public class ParserImpl<T> implements IParser<T> {

    private static final String TAG = "ParserImpl";

    private final Handler handler = new Handler(Looper.getMainLooper());

    @Override
    public ITabContext parseXSSFContext(Workbook workbook, Class<T> clazz) throws NotFindSheetException {
        String sheetName;
        Sheet sheet;
        TableContext tableContext = new TableContext();
        int numberOfSheets = workbook.getNumberOfSheets();
        Log.d(TAG, "parseContext: numberOfSheets number = " + numberOfSheets);
        ExcelTable annotation = clazz.getAnnotation(ExcelTable.class);
        if (annotation != null) {
            sheetName = annotation.sheetName();
            Log.d(TAG, "parseContext: sheetName = " + sheetName);
            sheet = workbook.getSheet(sheetName);
            tableContext.setSheetIndex(workbook.getSheetIndex(sheet));
            tableContext.setSheetName(sheetName);
            if (sheet == null) {
                throw new NotFindSheetException("not find sheet: " + sheetName);
            }
        } else {
            sheet = workbook.getSheetAt(0);
            sheetName = sheet.getSheetName();
            tableContext.setSheetIndex(0);
            tableContext.setSheetName(sheetName);
        }
        Row row = sheet.getRow(0);
        int physicalNumberOfCells = row.getPhysicalNumberOfCells();
        Log.d(TAG, "parseContext: physicalNumberOfCells number = " + physicalNumberOfCells);
        Map<Integer, String> headerMap = new HashMap<>();
        for (int columnIndex = 0; columnIndex < physicalNumberOfCells; columnIndex++) {
            Cell cell = row.getCell(columnIndex);
            String cellFormatValue = (String) getCellFormatValue(cell);
            headerMap.put(columnIndex, cellFormatValue);
        }
        tableContext.setExcelProperty(headerMap);
        Log.d(TAG, "parseContext: headerMap = " + headerMap);
        return tableContext;
    }

    @Override
    public void doParse(Workbook workbook, ITabContext context, Class<T> tClass, IParseListener<T> listener) {
        //获取表数
        int numberOfSheets = workbook.getNumberOfSheets();
        Log.d(TAG, "doXSSFParse: numberOfSheets number = " + numberOfSheets);
        Sheet sheetAtOne = workbook.getSheetAt(context.getSheetIndex());
        //获取行数
        int physicalNumberOfRows = sheetAtOne.getPhysicalNumberOfRows();
        Log.d(TAG, "doXSSFParse: physicalNumberOfRows number = " + physicalNumberOfRows);
        //跳过表头直接读取数据
        for (int row = 1; row < physicalNumberOfRows; row++) {
            parseData(sheetAtOne.getRow(row), context.getExcelProperty(), tClass, listener);
        }
        handler.post(listener::onEndParse);
    }

    private void parseData(Row rowData, Map<Integer, String> tabHeader, Class<T> tClass, IParseListener<T> listener) {
        try {
            Object item = tClass.newInstance();
            JSONArray extend = new JSONArray();
            for (int i = 0; i < rowData.getPhysicalNumberOfCells(); i++) {
                Cell cell = rowData.getCell(i);
                Object cellFormatValue = getCellFormatValue(cell);
                String headerName = tabHeader.get(i);
                boolean isMatch = false;
                for (Field declaredField : item.getClass().getDeclaredFields()) {
                    ExcelReadCell annotation = declaredField.getAnnotation(ExcelReadCell.class);
                    if (annotation != null) {
                        String name = annotation.name();
                        if (headerName != null && headerName.equals(name)) {
                            declaredField.setAccessible(true);
                            declaredField.set(item, cellFormatValue);
                            isMatch = true;
                            break;
                        }
                    }
                }
                if (!isMatch) {
                    JSONObject jsonObject = new JSONObject();
                    jsonObject.put("name", headerName);
                    jsonObject.put("value", cellFormatValue);
                    jsonObject.put("index", i);
                    extend.put(jsonObject);
                }
                for (Field declaredField : item.getClass().getDeclaredFields()) {
                    ExcelReadAggregate annotation = declaredField.getAnnotation(ExcelReadAggregate.class);
                    if (annotation != null) {
                        declaredField.setAccessible(true);
                        declaredField.set(item, extend.toString());
                    }
                }
            }
            handler.post(() -> listener.onParse((T) item, extend));
        } catch (InstantiationException | IllegalAccessException | JSONException e) {
            e.printStackTrace();
            handler.post(() -> listener.onParseError(e));
            Log.d(TAG, "parseData: ");
        }
    }


    private static Object getCellFormatValue(Cell cell) {
        @SuppressLint("SimpleDateFormat") SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");
        Object cellValue;
        if (cell != null) {
            switch (cell.getCellType()) {
                //文本，空白
                case Cell.CELL_TYPE_STRING:
                case Cell.CELL_TYPE_BLANK:
                    cellValue = cell.getStringCellValue();
                    break;
                //数字、日期
                case Cell.CELL_TYPE_NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        cellValue = fmt.format(cell.getDateCellValue());
                        //日期型
                    } else {
                        cellValue = String.valueOf(cell.getNumericCellValue());
                        //数字
                    }
                    break;
                //布尔型
                case Cell.CELL_TYPE_BOOLEAN:

                    cellValue = String.valueOf(cell.getBooleanCellValue());
                    break;
                //错误
                //公式
                case Cell.CELL_TYPE_FORMULA:
                    cellValue = "<公式类型无法解析>";
                    break;
                default:
                    cellValue = "";
            }

        } else {
            cellValue = null;
        }
        return cellValue;
    }
}
