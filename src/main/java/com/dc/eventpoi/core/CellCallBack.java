/**
 * RowStream.java
 */
package com.dc.eventpoi.core;

/**
 * @Description: excel单元格流式解析
 * @author 段超
 * @date: 2019年1月14日
 */
public interface CellCallBack extends BaseCallBack{
    /**
     * cell回调
     * @param excelcell cell对象
     * @author 段超
     * @date 2019-01-16 13:52:40
     */
    void getCell(ExcelCell excelcell);
}
