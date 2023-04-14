package com.dc.eventpoi.core.func;

import java.util.List;

import com.dc.eventpoi.core.inter.ExcelFunction;

/**
 * divideTruncate(age,100,2)
 * @author beijing-penguin
 *
 */
public class IfnullFunction implements ExcelFunction{

	@Override
	public String getName() {
		return "ifnull";
	}

	@Override
	public Object execute(List<Object> paramValueList) {
		Object k1 = paramValueList.get(0);
		Object k2 = paramValueList.get(1);
		if(k1 == null) {
			return k2;
		}else {
			return k1;
		}
	}

}
