package com.dc.eventpoi.core.entity;

import java.util.Arrays;
import java.util.List;

public class ListAndTableEntity {
	/**
	 * 表单集合
	 */
	private List<?> tableList;
	
	/**
	 * 数据集合
	 */
	private List<?> dataList;

	
	public static ListAndTableEntity build() {
		return new ListAndTableEntity();
	}

	public List<?> getTableList() {
		return tableList;
	}

	public ListAndTableEntity setTableList(List<?> tableList) {
		this.tableList = tableList;
		return this;
	}

	public ListAndTableEntity setTableList(Object tableObject) {
        this.tableList = Arrays.asList(tableObject);
        return this;
    }
	
	
	public List<?> getDataList() {
		return dataList;
	}

	public ListAndTableEntity setDataList(List<?> dataList) {
		this.dataList = Arrays.asList(dataList);
		return this;
	}
	
}
