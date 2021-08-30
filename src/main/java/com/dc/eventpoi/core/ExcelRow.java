/**
 * ExcelRow.java
 */
package com.dc.eventpoi.core;

import java.util.List;

/**
 * @Description: 行对象
 * @author beijing-penguin
 * @date: 2019年1月14日
 */
public class ExcelRow extends BaseExcelEntity{
    /**
     * 行索引
     */
    private int rowIndex;
    /**
     * 列集合
     */
    private List<ExcelCell> cellList;
    
    public int getRowIndex() {
        return rowIndex;
    }
    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }
    public List<ExcelCell> getCellList() {
        return cellList;
    }
    public void setCellList(List<ExcelCell> cellList) {
        this.cellList = cellList;
    }
}
