/**
 * ExcelStream.java
 */
package com.dc.eventpoi;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.List;

/**
 * @Description: excel大数据处理机遇事件流，行事件处理接口
 * @author 段超
 * @date: 2019年1月16日
 */
public interface ExcelEventStream {
	/**
	 * 指定sheet索引
	 * @param sheetIndexArr 索引集合
	 * @return ExcelEventStream
	 * @author 段超
	 * @date 2019-01-16 13:53:51
	 */
	ExcelEventStream sheetAt(Integer...sheetIndexArr);
	/**
	 * 得到事件发生时的工作簿名称
	 * @return String
	 * @author 段超
	 * @date 2019-01-16 13:54:20
	 */
	String getSheetName();
	/**
	 * 行回调方法
	 * @param baseCallBack 回调函数
	 * @throws Exception
	 * @author 段超
	 * @date 2019-01-16 13:54:51
	 */
	void rowStream(BaseCallBack baseCallBack) throws Exception;
	/**
	 * 得到事件结束后的工作簿集合
	 * @return List
	 * @author 段超
	 * @date 2019-01-16 13:55:04
	 */
	List<String> getSheetList();
	/**
	 * 得到事件发生时的工作簿索引
	 * @return short
	 * @author 段超
	 * @date 2019-01-16 13:55:25
	 */
	short getSheetIndex();
	/**
	 * 关闭所有流，清空对象内存
	 * @author 段超
	 * @date 2019-01-30 11:00:51
	 */
	void close();
	/**
	 * 读取文件
	 * @param file 文件
	 * @return ExcelEventStream
	 * @author 段超
	 * @throws Exception 
	 * @date 2019-01-16 13:55:44
	 */
	static ExcelEventStream readExcel(File file) throws Exception {
		return readExcel(new FileInputStream(file));
	}
	/**
	 * 读取excel
	 * @param bytes 文件二进制数据
	 * @return ExcelEventStream
	 * @author 段超
	 * @throws Exception 
	 * @date 2019-01-16 13:55:44
	 */
	static ExcelEventStream readExcel(byte[] bytes) throws Exception {
		return readExcel(new ByteArrayInputStream(bytes));
	}
	/**
	 * 读取文件
	 * @param fileStream 文件流
	 * @return ExcelEventStream
	 * @author 段超
	 * @date 2019-01-16 13:55:44
	 */
	static ExcelEventStream readExcel(InputStream fileStream) throws Exception {
		FileType fileType = ExcelHelper.judgeFileType(fileStream);
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
