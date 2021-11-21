package com.dc.eventpoi.test;

import java.io.InputStream;

import com.alibaba.fastjson.JSON;
import com.dc.eventpoi.core.entity.ExcelRow;
import com.dc.eventpoi.core.inter.ExcelEventStream;
import com.dc.eventpoi.core.inter.RowCallBack;

/**
 * 按行回调的方式读取数据，回调返回 ExcelRow 对象
 * @author DC
 *
 */
public class ExcelEventStreamTest {
	public static void main(String[] args) {
		try {
			//行回调
			InputStream excelInput = ExcelEventStreamTest.class.getResourceAsStream("demo1.xlsx");
			ExcelEventStream stream = ExcelEventStream.readExcel(excelInput);
			stream.rowStream(new RowCallBack() {
				@Override
				public void getRow(ExcelRow row) {
					System.out.println(stream.getSheetName() +"=="+JSON.toJSONString(row));
				}
			});
			stream.close();
		}catch (Exception e) {
			e.printStackTrace();
		}
	}
}
