/**
 * BaseExcelStream.java
 */
package com.dc.eventpoi.core;

import com.dc.eventpoi.core.entity.BaseExcelEntity;
import com.dc.eventpoi.core.entity.ExcelRow;
import com.dc.eventpoi.core.inter.BaseCallBack;
import com.dc.eventpoi.core.inter.RowCallBack;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 基础excel流对象
 *
 * @author beijing-penguin
 */
public class BaseExcelStream {
    /**
     * 指定sheet数组
     */
    protected Integer[] sheetIndexArr;
    /**
     * sheet集合
     */
    protected List<String> sheetList = new ArrayList<String>(10);
    /**
     * 文件流
     */
    protected InputStream fileStream;

    /**
     * @param baseCallBack 回调
     * @param entity       实体
     */
    protected void excuteCallBack(BaseCallBack baseCallBack, BaseExcelEntity entity) {
        if (BaseCallBack.class.isAssignableFrom(RowCallBack.class)) {
            ((RowCallBack) baseCallBack).getRow((ExcelRow) entity);
        } else if (true) {
            // TODO 开发cell单元格回调
        }
    }


    public List<String> getSheetList() {
        return sheetList;
    }
}
