package com.dc.eventpoi.core;

import java.math.BigDecimal;
import java.util.Map;

import com.googlecode.aviator.runtime.function.AbstractFunction;
import com.googlecode.aviator.runtime.type.AviatorObject;
import com.googlecode.aviator.runtime.type.AviatorString;

public class RmZeroFunction extends AbstractFunction {
	
	private static final long serialVersionUID = 1L;

	@Override
    public String getName() {
        return "rmzero";
    }

    @Override
    public AviatorString call(Map<String, Object> env, AviatorObject arg1) {
    	Object v1 = arg1.getValue(env);
        if(v1 != null) {
        	return new AviatorString(new BigDecimal(v1.toString()).stripTrailingZeros().toPlainString());
        }else {
        	return new AviatorString(null);
        }
    }
}