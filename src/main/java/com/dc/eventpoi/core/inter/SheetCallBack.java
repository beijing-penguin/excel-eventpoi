package com.dc.eventpoi.core.inter;

import org.apache.poi.xssf.streaming.SXSSFSheet;

/**
 * 行回调接口
 *
 * @author beijing-penguin
 */
public interface SheetCallBack extends BaseCallBack {

    /**
     * 行回调方法
     *
     * @param excelSheet 行对象
     */
    void callBack(SXSSFSheet sxssSheet);
}
