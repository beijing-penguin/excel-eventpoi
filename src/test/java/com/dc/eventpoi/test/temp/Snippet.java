package com.dc.eventpoi.test.temp;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import com.googlecode.aviator.AviatorEvaluator;

public class Snippet {
	public static void main(String[] args) {
		Map<String, Object> dataMap = new HashMap<>();
	     dataMap.put("x", 1);
	     dataMap.put("y", 2);
	     dataMap.put("z", 3);
	     String expression = "string.substring(str((x + y) * x + (x + z) * y /0.7),0,2)";
	     
	     
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

