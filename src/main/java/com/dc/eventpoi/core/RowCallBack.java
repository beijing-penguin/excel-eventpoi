/**
 * RowCallBack.java
 */
package com.dc.eventpoi.core;

/**
 * @Description: 行回调接口
 * @author 段超
 * @date: 2019年1月14日
 */
public interface RowCallBack extends BaseCallBack{

    /**
     * 行回调方法
     * @param row 行对象
     * @author 段超
     * @date 2019-01-16 14:04:45
     */
    void getRow(ExcelRow row);
}
