package com.dc.eventpoi.core.entity;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class ListAndTableEntity {
	/**
	 * 表单集合
	 */
	private List<Map<?, ?>> tableList;
	
	/**
	 * 数据集合
	 */
	private List<List<Map<?, ?>>> dataList = new ArrayList<>();

	
	public static ListAndTableEntity build() {
		return new ListAndTableEntity();
	}

	public List<?> getTableList() {
		return tableList;
	}

	public ListAndTableEntity setTableList(List<?> tableList) throws Throwable {
		List<Map<?, ?>> tableMapList = new ArrayList<Map<?, ?>>();
		for(Object obj : tableList) {
			if(obj.getClass().getClassLoader() == null) {
				LinkedHashMap<String, Object> map = new LinkedHashMap<>();
				Field[] v_obj_field_arr = obj.getClass().getDeclaredFields();
				for (Field field : v_obj_field_arr) {
			        field.setAccessible(true);
			        String keyName = field.getName();
			        map.put(keyName, field.get(obj));
			    }
				tableMapList.add(map);
			}else {
				tableMapList.add((Map<?, ?>)obj);
			}
		}
		this.tableList = tableMapList;
		return this;
	}

	
	public List<List<Map<?, ?>>> getDataList() {
		return dataList;
	}

	public ListAndTableEntity setDataList(List<?> dataList) {
		for(Object obj : dataList) {
			List<?> objList = (List<?>)obj;
			for(Object oo : objList) {
				
			}
			List<Map<?, ?>> newMapList = (List<?>)obj;
		}
		return this;
	}
	
}
