/**
 * ExcelRow.java
 */
package com.dc.eventpoi.core.entity;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;

/**
 * 单元格实体对象
 *
 * @author beijing-penguin
 */
public class ExportExcelCell {
    /**
     * 列索引
     */
    private Short index;
    /**
     * 值
     */
    private String value;

    /**
     * cellStyle
     */
    private CellStyle cellStyle;
    /**
     * cellType
     */
    private CellType cellType;

    /**
     * @param index 列索引
     * @param value 值
     * @param cellStyle 单元格样式
     */
    public ExportExcelCell(Short index, String value, CellStyle cellStyle) {
        this.index = index;
        this.value = value;
        this.cellStyle = cellStyle;
    }

    public Short getIndex() {
        return index;
    }


    public void setIndex(Short index) {
        this.index = index;
    }


    public String getValue() {
        return value;
    }


    public void setValue(String value) {
        this.value = value;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public CellType getCellType() {
        return cellType;
    }

    public void setCellType(CellType cellType) {
        this.cellType = cellType;
    }
    
}
