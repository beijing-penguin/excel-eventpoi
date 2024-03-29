package com.dc.eventpoi.core;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

public class ExportUtils {

	private static String LETTER_DIGIT_REGEX = "^[a-z0-9A-Z]+$";

	public static boolean isAZAndDigit(String str) {
		return str.matches(LETTER_DIGIT_REGEX);
	}

	public static boolean isWordKeyPart(String str) {
		return !str.matches(LETTER_DIGIT_REGEX);
	}

	public static String stripTrailingZeros(String str) {
		try {
			BigDecimal ss = new BigDecimal(str);
			return ss.stripTrailingZeros().toPlainString();
		}catch (Exception e) {
			return str;
		}
   }
	
	public static List<String> expOnlyKeyStr(String str,String...targetArr) {
		List<String> expKeyList = new ArrayList<>();
		
		boolean isSimpleExp = true;
		for(String target : targetArr) {
			if(isSimpleExp == false) {
				break;
			}
			int count = 0;
		    int index = 0;
		    while ((index = str.indexOf(target, index)) != -1) {
		        count++;
		        index += target.length();
		    }
		    if(count == 1) {
		    	expKeyList.add(target);
		    }else if(count > 1){
		    	isSimpleExp = false;
		    }
		}
		if(isSimpleExp == true ) {
			return expKeyList;
		}else {
			return null;
		}
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
				if(!keyList.contains(addStr)) {
					keyList.add(addStr);
				}
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
}
