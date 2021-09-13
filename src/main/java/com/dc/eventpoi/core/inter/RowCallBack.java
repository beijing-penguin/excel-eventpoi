/**
 * RowCallBack.java
 */
package com.dc.eventpoi.core.inter;

import com.dc.eventpoi.core.entity.ExcelRow;

/**
 * 行回调接口
 *
 * @author beijing-penguin
 */
public interface RowCallBack extends BaseCallBack {

    /**
     * 行回调方法
     *
     * @param row 行对象
     */
    void getRow(ExcelRow row);
}
