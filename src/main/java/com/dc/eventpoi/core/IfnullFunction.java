//package com.dc.eventpoi.core;
//
//import java.util.Map;
//
//import com.googlecode.aviator.runtime.function.AbstractFunction;
//import com.googlecode.aviator.runtime.type.AviatorObject;
//import com.googlecode.aviator.runtime.type.AviatorString;
//
//public class IfnullFunction extends AbstractFunction {
//	
//	private static final long serialVersionUID = 1L;
//
//	@Override
//    public String getName() {
//        return "ifnull";
//    }
//
//    @Override
//    public AviatorString call(Map<String, Object> env, AviatorObject arg1, AviatorObject arg2) {
//    	Object v1 = arg1.getValue(env);
//    	Object v2 = arg2.getValue(env);
//        if(v1 != null) {
//        	return new AviatorString(v1.toString());
//        }else {
//        	if(v2 != null) {
//        		return new AviatorString(arg2.getValue(env).toString());
//        	}else {
//        		return null;
//        	}
//        }
//    }
//}