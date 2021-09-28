package com.dc.eventpoi.core.inter;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFCell;

public interface CallBackCellStyle {
    /**
     * 单元格样式回调
     * @param curCell 当前单元格对象
     * @param cellStyle 单元格样式
     */
    void callBack(SXSSFCell curCell,CellStyle cellStyle);
}
