package com.dc.eventpoi.core;

import java.util.List;

/**
 * 工作簿对象
 * @author beijing-penguin
 * @date: 2019年2月18日
 */
public class ExcelSheet extends BaseExcelEntity{
	/**
	 * 
	 */
	private Integer sheetIndex;
	/**
	 * 
	 */
	private String sheetName;
	/**
	 * 
	 */
	private List<ExcelRow> rowList;
	
	public Integer getSheetIndex() {
		return sheetIndex;
	}
	public void setSheetIndex(Integer sheetIndex) {
		this.sheetIndex = sheetIndex;
	}
	public String getSheetName() {
		return sheetName;
	}
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	public List<ExcelRow> getRowList() {
		return rowList;
	}
	public void setRowList(List<ExcelRow> rowList) {
		this.rowList = rowList;
	}
	
	
}
