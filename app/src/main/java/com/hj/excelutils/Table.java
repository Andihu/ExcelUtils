package com.hj.excelutils;

import com.hj.excellibrary.annotation.ExcelReadAggregate;
import com.hj.excellibrary.annotation.ExcelReadCell;
import com.hj.excellibrary.annotation.ExcelTable;
import com.hj.excellibrary.annotation.ExcelWriteAdapter;
import com.hj.excellibrary.annotation.ExcelWriteCell;

@ExcelTable(sheetName = "测试表1")
public class Table {

    @ExcelReadCell(name = "规格")
    @ExcelWriteCell(writeIndex = 3, writeName = "规格")
    public String specification;

    @ExcelReadCell(name = "责任部门")
    @ExcelWriteCell(writeIndex = 0, writeName = "责任部门")
    public String department;

    @ExcelReadCell(name = "责任人")
    @ExcelWriteCell(writeIndex = 5, writeName = "责任人")
    public String responsiblePerson;

    @ExcelReadCell(name = "存放位置")
    @ExcelWriteCell(writeIndex = 4, writeName = "存放位置")
    public String storageLocation;

    @ExcelReadCell(name = "备注")
    @ExcelWriteCell(writeIndex = 6, writeName = "备注")
    public String note;

    @ExcelReadCell(name = "物品名称")
    @ExcelWriteCell(writeIndex = 2, writeName = "物品名称")
    public String name;

    @ExcelReadCell(name = "物品编码")
    @ExcelWriteCell(writeIndex = 1, writeName = "物品编码")
    public String code;

    @ExcelReadAggregate
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
