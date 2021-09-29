package com.dc.eventpoi.core.inter;

import org.apache.poi.xssf.streaming.SXSSFSheet;

/**
 * 行回调接口
 *
 * @author beijing-penguin
 */
public interface SheetCallBack extends BaseCallBack {

    /**
     * Sheet回调
     * @param sxssSheet sxssSheet
     */
    void callBack(SXSSFSheet sxssSheet);
}
