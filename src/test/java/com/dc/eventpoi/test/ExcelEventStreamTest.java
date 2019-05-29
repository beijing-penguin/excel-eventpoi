package com.dc.eventpoi.test;

import java.io.InputStream;

import com.alibaba.fastjson.JSON;
import com.dc.eventpoi.ExcelEventStream;
import com.dc.eventpoi.ExcelRow;
import com.dc.eventpoi.RowCallBack;

public class ExcelEventStreamTest {
	public static void main(String[] args) {
		try {
			//行回调
			InputStream excelInput = Test1.class.getResourceAsStream("demo1.xlsx");
			ExcelEventStream stream = ExcelEventStream.readExcel(excelInput);
			stream.rowStream(new RowCallBack() {
				@Override
				public void getRow(ExcelRow row) {
					System.out.println(JSON.toJSONString(row));
				}
			});
			stream.close();
		}catch (Exception e) {
			e.printStackTrace();
		}
	}
}
