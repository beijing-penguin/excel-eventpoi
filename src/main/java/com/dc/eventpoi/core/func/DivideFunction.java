package com.dc.eventpoi.core.func;

import java.util.List;

import com.dc.eventpoi.core.inter.ExcelFunction;

/**
 * divideRound(age,100)
 * @author beijing-penguin
 *
 */
public class DivideFunction implements ExcelFunction{

	@Override
	public String getName() {
		return "divide";
	}

	@Override
	public Object execute(List<Object> paramValueList) {
		Object k1 = paramValueList.get(0);
		Object k2 = paramValueList.get(1);
		return Double.valueOf(k1.toString())/Double.valueOf(k2.toString());
	}

}
