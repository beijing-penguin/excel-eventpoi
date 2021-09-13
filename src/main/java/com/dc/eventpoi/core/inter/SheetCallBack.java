package com.dc.eventpoi.core.inter;

import com.dc.eventpoi.core.entity.ExcelSheet;

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
    void getSheet(ExcelSheet excelSheet);
}
