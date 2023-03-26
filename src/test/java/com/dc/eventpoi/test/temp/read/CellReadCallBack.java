package com.dc.eventpoi.test.temp.read;

public class CellReadCallBack implements StreamReadBaseCallBack{
	
	private short cellIndex;
	
	private int rowIndex;
	
	private String cellNo;
	
	private String cellValue;

	
	public short getCellIndex() {
		return cellIndex;
	}

	public void setCellIndex(short cellIndex) {
		this.cellIndex = cellIndex;
	}

	public int getRowIndex() {
		return rowIndex;
	}

	public void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}

	public String getCellNo() {
		return cellNo;
	}

	public void setCellNo(String cellNo) {
		this.cellNo = cellNo;
	}

	public String getCellValue() {
		return cellValue;
	}

	public void setCellValue(String cellValue) {
		this.cellValue = cellValue;
	}


	
}
