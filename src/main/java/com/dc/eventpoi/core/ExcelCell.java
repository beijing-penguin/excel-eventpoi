/**
 * ExcelRow.java
 */
package com.dc.eventpoi.core;

/**
 * @Description: 单元格实体对象
 * @author beijing-penguin
 * @date: 2019年1月14日
 */
public class ExcelCell extends BaseExcelEntity{
    /**
     * 列索引
     */
    private Short index;
    /**
     * 值
     */
    private String value;
    /**
     * 图片数组
     */
    private byte[] imgBytes;
    
    /**
     * 
     * @param index 列索引
     * @param value 值
     */
    public ExcelCell(Short index, String value) {
        this.index = index;
        this.value = value;
    }

    /**
     * 
     * @param index 列索引
     * @param value 值
     */
    public ExcelCell(Short index, byte[] value) {
        this.index = index;
        this.imgBytes = value;
    }
    
    public Short getIndex() {
        return index;
    }


    public void setIndex(Short index) {
        this.index = index;
    }


    public String getValue() {
        return value;
    }


    public void setValue(String value) {
        this.value = value;
    }


    public byte[] getImgBytes() {
        return imgBytes;
    }


    public void setImgBytes(byte[] imgBytes) {
        this.imgBytes = imgBytes;
    }
    
}
