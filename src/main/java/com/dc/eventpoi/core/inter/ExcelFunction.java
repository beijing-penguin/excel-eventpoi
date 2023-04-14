package com.dc.eventpoi.core.inter;

import java.util.List;

public interface ExcelFunction {
	String getName();

	Object execute(List<Object> paramValueList);
}
