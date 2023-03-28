package com.dc.eventpoi.test.temp;

import java.util.List;

import com.alibaba.fastjson.JSON;
import com.dc.eventpoi.test.temp.read.CellReadCallBack;
import com.dc.eventpoi.test.temp.read.XlsxReadStream;


public class TestReadXlsx {
	public static void main(String[] args) throws Throwable  {
    	XlsxReadStream howto = new XlsxReadStream();
    	howto.setReadSheetIndex(0);
    	howto.setFileName("E:\\eclipse-workspace-2022-12\\excel-eventpoi\\my_test_temp\\file.xlsx");
    	howto.doRead(new RowCallBack() {
			@Override
			public void callBack(int rowIndex, List<CellReadCallBack> cellList) {
				System.err.println("rowIndex="+rowIndex+",list="+JSON.toJSONString(cellList));
			}
		});
    	howto.doRead(new CellCallBack() {
			@Override
			public void callBack(CellReadCallBack cell) {
				System.err.println("cell="+JSON.toJSONString(cell));
			}
		});
    	System.err.println(JSON.toJSONString(howto.doRead()));
    }
	
}
