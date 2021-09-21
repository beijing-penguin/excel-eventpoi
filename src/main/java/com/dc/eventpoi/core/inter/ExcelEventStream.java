/**
 * ExcelStream.java
 */
package com.dc.eventpoi.core.inter;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.List;

import com.dc.eventpoi.core.ExcelXlsStream;
import com.dc.eventpoi.core.ExcelXlsxStream;
import com.dc.eventpoi.core.PoiUtils;
import com.dc.eventpoi.core.enums.FileType;

/**
 * excel大数据处理机遇事件流，行事件处理接口
 *
 * @author beijing-penguin
 */
public interface ExcelEventStream {
    /**
     * 指定sheet索引
     *
     * @param sheetIndexArr 索引集合
     * @return ExcelEventStream
     */
    ExcelEventStream sheetAt(Integer... sheetIndexArr);

    /**
     * 得到事件发生时的工作簿名称
     *
     * @return String
     */
    String getSheetName();

    /**
     * 行回调方法
     *
     * @param baseCallBack 回调函数
     * @throws Exception 回调异常
     */
    void rowStream(BaseCallBack baseCallBack) throws Exception;

    /**
     * 得到事件结束后的工作簿集合
     *
     * @return List
     */
    List<String> getSheetList();

    /**
     * 得到事件发生时的工作簿索引
     *
     * @return short
     */
    short getSheetIndex();

    /**
     * 关闭所有流，清空对象内存
     *
     * @throws Exception IOException
     */
    void close() throws Exception;

    /**
     * 读取文件
     *
     * @param file 文件
     * @return ExcelEventStream
     * @throws Exception IOException
     */
    static ExcelEventStream readExcel(File file) throws Exception {
        return readExcel(new FileInputStream(file));
    }

    /**
     * 读取excel
     *
     * @param bytes 文件二进制数据
     * @return ExcelEventStream
     * @throws Exception IOException
     */
    static ExcelEventStream readExcel(byte[] bytes) throws Exception {
        return readExcel(new ByteArrayInputStream(bytes));
    }


    /**
     * 读取文件
     *
     * @param fileStream 文件流
     * @return ExcelEventStream
     * @throws Exception IOException
     */
    static ExcelEventStream readExcel(InputStream fileStream) throws Exception {
        FileType fileType = PoiUtils.judgeFileType(fileStream);
        switch (fileType) {
            case XLS:
                return new ExcelXlsStream(fileStream);
            case XLSX:
                return new ExcelXlsxStream(fileStream);
            default:
                throw new Exception("filetype is unsupport");
        }
    }
}
