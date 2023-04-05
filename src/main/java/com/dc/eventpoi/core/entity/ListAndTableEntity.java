package com.dc.eventpoi.core.entity;

import java.util.ArrayList;
import java.util.List;

public class ListAndTableEntity {
	/**
	 * 表单集合
	 */
	private List<Object> tableList;
	
	/**
	 * 数据集合
	 */
	private List<Object> dataList;

	
	public static ListAndTableEntity build() {
		return new ListAndTableEntity();
	}

	public List<?> getTableList() {
		return tableList;
	}

	public ListAndTableEntity setTable(Object tableObject) {
		if(tableList == null) {
			tableList = new ArrayList<>();
		}
		tableList.add(tableObject);
        return this;
    }
	
	
	public List<?> getDataList() {
		return dataList;
	}

	public ListAndTableEntity setDataList(List<?> dataList) {
		if(this.dataList == null) {
			this.dataList = new ArrayList<>();
		}
		this.dataList.add(dataList);
		return this;
	}
	
}
