package com.dc.eventpoi.test.temp;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import com.googlecode.aviator.AviatorEvaluator;

public class Snippet2 {
	public static void main(String[] args) {
		Map<String, Object> dataMap = new HashMap<>();
	     dataMap.put("list.x", 1);
	     dataMap.put("y", 2);
	     dataMap.put("z", null);
	     dataMap.put("zzzz", 3);
	     String expression = "z";
	     
	     
	     long t2 = System.currentTimeMillis();
	     for (int i = 0; i < 10; i++) {
	    	 System.err.println(expression+"="+ AviatorEvaluator.compile(expression).execute(dataMap));
	     }
	     System.err.println(System.currentTimeMillis()-t2);
	     
	     long t1 = System.currentTimeMillis();
	     for (int i = 0; i < 10; i++) {
		     System.err.println(expression+"="+ AviatorEvaluator.execute(expression, dataMap));
	     }
	     
	     System.err.println(System.currentTimeMillis()-t1);
	     
	}
}

