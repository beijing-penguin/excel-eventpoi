package com.dc.eventpoi.test.temp.write;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import com.alibaba.fastjson.JSON;
import com.dc.eventpoi.test.temp.read.CellReadCallBack;

public class ExportUtils {

	private static String LETTER_DIGIT_REGEX = "^[a-z0-9A-Z]+$";
	public static void main(String[] args) {
		System.err.println(JSON.toJSONString(getExpAllKeys("${1-a *2}%${list.headImage/tab.age}")));
	}

	public static boolean isAZAndDigit(String str) {
		return str.matches(LETTER_DIGIT_REGEX);
	}

	public static boolean isWordKeyPart(String str) {
		return !str.matches(LETTER_DIGIT_REGEX);
	}

	public static List<String> getExpAllKeys(String keyStr){
		List<String> keyList = new ArrayList<>();


		String tempStr = "";
		for (int i = 0; i < keyStr.length(); i++) {
			tempStr = tempStr + keyStr.charAt(i);
			String addStr = null;
			if(tempStr.contains("list.")) {
				addStr = tempStr.substring(tempStr.indexOf("list."));
			}
			if(tempStr.contains("tab.")) {
				addStr = tempStr.substring(tempStr.indexOf("tab."));
			}
			if(addStr != null) {
				//往后查找不是字母或者数字
				for (int j = i+1; j < keyStr.length(); j++) {
					i = j-1;
					if(String.valueOf(keyStr.charAt(j)).matches(LETTER_DIGIT_REGEX)) {
						addStr = addStr + keyStr.charAt(j);
					}else {
						break;
					}
				}
				keyList.add(addStr);
				tempStr = "";
			}
		}

		return keyList;
	}

	public static void setExpMap(Object tabOrOneListObj,String cell_key,Map<String, Object> expMap) throws Throwable {
		if(tabOrOneListObj instanceof Map) {
			Map<?,?> v_map = (Map<?,?>)tabOrOneListObj;
			if(v_map.size() > 0 ) {
				for(Entry<?,?> entry : v_map.entrySet()) {
					String keyName = entry.getKey().toString();
					String keyName_word = "tab."+keyName;
					if(cell_key.contains(keyName_word)) {
						Object value = entry.getValue();
						if (value != null) {
							expMap.put(keyName_word, value);
						}
					}
				}
			}
		}else {
			Field[] fields = tabOrOneListObj.getClass().getDeclaredFields();
			for (Field field : fields) {
				field.setAccessible(true);
				String keyName = field.getName();
				String keyName_word = "list."+keyName;
				if(cell_key.contains(keyName_word) ) {
					Object value = field.get(tabOrOneListObj);
					if (value != null) {
						expMap.put(keyName_word, value);
					}
				}
			}
		}
	}

	public static int findListIndexByKey(List<CellReadCallBack> tempContentCollection,String cur_cell_value,int curRowIndex) {
		int keyRowIndex = -1;
		for (CellReadCallBack cell:tempContentCollection) {
			if(cell.getCellValue().equals(cur_cell_value)) {
				keyRowIndex = cell.getRowIndex();
			}
		}
		return curRowIndex-keyRowIndex;
	} 
}
