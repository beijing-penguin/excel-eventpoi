package com.dc.eventpoi.core.func;

import java.math.BigDecimal;
import java.util.List;

import com.dc.eventpoi.core.inter.ExcelFunction;

/**
 * divideRound(age,100,2)
 * @author beijing-penguin
 *
 */
public class DivideRoundFunction implements ExcelFunction{

	@Override
	public String getName() {
		return "divideRound";
	}

	@Override
	public Object execute(List<Object> paramValueList) {
		Object k1 = paramValueList.get(0);
		Object k2 = paramValueList.get(1);
		Object k3 = paramValueList.get(2);
		return new BigDecimal(k1.toString()).divide(new BigDecimal(k2.toString()),Integer.valueOf(k3.toString()),BigDecimal.ROUND_HALF_UP).stripTrailingZeros().toPlainString();
	}

}
