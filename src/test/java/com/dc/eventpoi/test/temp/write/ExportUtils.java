package com.dc.eventpoi.test.temp.write;

import java.lang.reflect.Field;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import com.dc.eventpoi.test.temp.read.CellReadCallBack;

public class ExportUtils {
	
	private static String LETTER_DIGIT_REGEX = "^[a-z0-9A-Z]+$";
	public static void main(String[] args) {
		System.err.println(" ".matches(LETTER_DIGIT_REGEX));
	}
	
	public static boolean isWordKeyPart(String str) {
		return !str.matches(LETTER_DIGIT_REGEX);
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
	
	public static int findListIndexByKey(Collection<List<CellReadCallBack>> tempContentCollection,String cur_cell_value,int curRowIndex) {
		Iterator<List<CellReadCallBack>> cellIt = tempContentCollection.iterator();
		int keyRowIndex = -1;
		while(cellIt.hasNext()) {
			List<CellReadCallBack> cellList = cellIt.next();
			for (CellReadCallBack cell:cellList) {
				if(cell.getCellValue().equals(cur_cell_value)) {
					keyRowIndex = cell.getRowIndex();
				}
			}
		}
		return keyRowIndex-curRowIndex;
	} 
}
