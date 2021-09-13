/**
 * RowStream.java
 */
package com.dc.eventpoi.core.inter;

import com.dc.eventpoi.core.entity.ExcelCell;

/**
 * excel单元格流式解析
 *
 * @author beijing-penguin
 */
public interface CellCallBack extends BaseCallBack {
    /**
     * cell回调
     *
     * @param excelcell cell对象
     */
    void getCell(ExcelCell excelcell);
}
