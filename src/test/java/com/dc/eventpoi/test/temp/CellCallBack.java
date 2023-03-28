package com.dc.eventpoi.test.temp;

import com.dc.eventpoi.test.temp.read.CellReadCallBack;

public interface CellCallBack extends RegCallBack {
	public void callBack(CellReadCallBack cell);
}
