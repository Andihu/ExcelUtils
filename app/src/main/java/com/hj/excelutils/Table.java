package com.hj.excelutils;

import com.hj.excellibrary.annotation.CellStyle;
import com.hj.excellibrary.annotation.ExcelReadAggregate;
import com.hj.excellibrary.annotation.ExcelReadCell;
import com.hj.excellibrary.annotation.ExcelTable;
import com.hj.excellibrary.annotation.ExcelWriteAdapter;
import com.hj.excellibrary.annotation.ExcelWriteCell;

import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

@ExcelTable(sheetName = "测试表1")
public class Table {

    @ExcelReadCell(name = "规格")
    @ExcelWriteCell(writeIndex = 3, writeName = "规格")
    @CellStyle(horizontalAlign = HorizontalAlignment.CENTER, verticalAlign = VerticalAlignment.CENTER, underline = FontUnderline.DOUBLE)
    public String specification;

    @ExcelReadCell(name = "责任部门")
    @ExcelWriteCell(writeIndex = 0, writeName = "责任部门")
    @CellStyle(horizontalAlign = HorizontalAlignment.CENTER, verticalAlign = VerticalAlignment.CENTER, underline = FontUnderline.DOUBLE)
    public String department;

    @ExcelReadCell(name = "责任人")
    @ExcelWriteCell(writeIndex = 5, writeName = "责任人")
    @CellStyle(horizontalAlign = HorizontalAlignment.CENTER, verticalAlign = VerticalAlignment.CENTER, underline = FontUnderline.DOUBLE)
    public String responsiblePerson;

    @ExcelReadCell(name = "存放位置")
    @ExcelWriteCell(writeIndex = 4, writeName = "存放位置")
    @CellStyle(horizontalAlign = HorizontalAlignment.CENTER, verticalAlign = VerticalAlignment.CENTER, underline = FontUnderline.DOUBLE)
    public String storageLocation;

    @ExcelReadCell(name = "备注")
    @ExcelWriteCell(writeIndex = 6, writeName = "备注")
    @CellStyle(horizontalAlign = HorizontalAlignment.CENTER, verticalAlign = VerticalAlignment.CENTER, underline = FontUnderline.DOUBLE)
    public String note;

    @ExcelReadCell(name = "物品名称")
    @ExcelWriteCell(writeIndex = 2, writeName = "物品名称")
    @CellStyle(horizontalAlign = HorizontalAlignment.CENTER, verticalAlign = VerticalAlignment.CENTER, underline = FontUnderline.SINGLE)
    public String name;

    @ExcelReadCell(name = "物品编码")
    @ExcelWriteCell(writeIndex = 1, writeName = "物品编码")
    @CellStyle(horizontalAlign = HorizontalAlignment.CENTER, verticalAlign = VerticalAlignment.CENTER, underline = FontUnderline.SINGLE)
    public String code;

    @ExcelReadAggregate
    @CellStyle(horizontalAlign = HorizontalAlignment.CENTER, verticalAlign = VerticalAlignment.CENTER, underline = FontUnderline.SINGLE)
    @ExcelWriteAdapter(adapter = JsonArrayConvertAdapter.class)
    public String extend;

    @Override
    public String toString() {
        return "Test{" +
                "specification='" + specification + '\'' +
                ", department='" + department + '\'' +
                ", responsiblePerson='" + responsiblePerson + '\'' +
                ", storageLocation='" + storageLocation + '\'' +
                ", note='" + note + '\'' +
                ", name='" + name + '\'' +
                ", code='" + code + '\'' +
                ", extend='" + extend + '\'' +
                '}';
    }
}
