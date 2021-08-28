/**
 * BaseExcelStream.java
 */
package com.dc.eventpoi.core;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @Description: 基础excel流对象
 * @author 段超
 * @date: 2019年1月18日
 */
public class BaseExcelStream{
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
     * @param entity 实体
     * @author 段超
     * @date 2019-01-18 10:36:42
     */
    protected void excuteCallBack(BaseCallBack baseCallBack,BaseExcelEntity entity) {
        if(BaseCallBack.class.isAssignableFrom(RowCallBack.class)) {
            ((RowCallBack)baseCallBack).getRow((ExcelRow)entity);
        }else if(true) {
            // TODO 开发cell单元格回调
        }
    }
    
    
    public List<String> getSheetList() {
        return sheetList;
    }
}
