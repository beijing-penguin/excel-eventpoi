package com.dc.eventpoi.core.inter;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFSheet;

public interface CellStyleCallBack {
    /**
     * 单元格样式回调
     * @param sxssSheet 当前sxssSheet
     * @param curCell 当前单元格对象
     * @param curCellStyle 当前单元格样式
     */
    void callBack(SXSSFSheet sxssSheet,SXSSFCell curCell,CellStyle curCellStyle);
}
