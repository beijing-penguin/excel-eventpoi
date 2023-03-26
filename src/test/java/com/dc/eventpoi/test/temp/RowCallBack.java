package com.dc.eventpoi.test.temp;

import java.util.List;

import com.dc.eventpoi.test.temp.read.CellReadCallBack;

public interface RowCallBack extends RegCallBack{
	public void callBack(int rowIndex, List<CellReadCallBack> cellList);
}
