/**
 * RowStream.java
 */
package com.dc.eventpoi.core.inter;

import com.dc.eventpoi.core.entity.ExcelCell;

/**
 * @Description: excel单元格流式解析
 * @author beijing-penguin
 * @date: 2019年1月14日
 */
public interface CellCallBack extends BaseCallBack{
    /**
     * cell回调
     * @param excelcell cell对象
     * @date 2019-01-16 13:52:40
     */
    void getCell(ExcelCell excelcell);
}
