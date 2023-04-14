//package com.dc.eventpoi.core;
//
//import java.math.BigDecimal;
//import java.util.Map;
//
//import com.googlecode.aviator.runtime.function.AbstractFunction;
//import com.googlecode.aviator.runtime.type.AviatorObject;
//import com.googlecode.aviator.runtime.type.AviatorString;
//
//public class RoundFunction extends AbstractFunction {
//	
//	private static final long serialVersionUID = 1L;
//
//	@Override
//    public String getName() {
//        return "round";
//    }
//
//	@Override
//    public AviatorObject call(Map<String, Object> env, AviatorObject arg1) {
//        if (arg1 == null) {
//            throw new NullPointerException("round function args can't be null");
//        }
//        double num = Double.parseDouble(arg1.getValue(env).toString());
//        int scale = 0;
//        BigDecimal bd = new BigDecimal(num);
//        bd = bd.setScale(scale, BigDecimal.ROUND_HALF_UP);
//        return new AviatorString(bd.toPlainString());
//    }
//	
//    @Override
//    public AviatorObject call(Map<String, Object> env, AviatorObject arg1, AviatorObject arg2) {
//        if (arg1 == null) {
//            throw new NullPointerException("round function args can't be null");
//        }
//        int scale = 0;
//        if(arg2 != null) {
//        	scale = Integer.parseInt(arg2.getValue(env).toString());
//        }
//        
//        BigDecimal bd = new BigDecimal(arg1.getValue(env).toString());
//        bd = bd.setScale(scale, BigDecimal.ROUND_HALF_UP);
//        return new AviatorString(bd.toPlainString());
//    }
//}