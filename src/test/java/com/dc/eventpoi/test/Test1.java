package com.dc.eventpoi.test;

import java.io.InputStream;
import java.util.List;

import com.alibaba.fastjson.JSON;
import com.dc.eventpoi.ExcelHelper;
import com.dc.eventpoi.ExcelRow;

public class Test1 {
	public static void main(String[] args) {
		/***********返回List<ExcelRow>类型数据*********************************/
		InputStream excelInput = Test1.class.getResourceAsStream("demo1.xlsx");
		System.out.println(excelInput);
		try {
			List<ExcelRow> dataList1 = ExcelHelper.parseExcelRowList(excelInput);//默认只读取所有工作簿数据，第二个参数指定工作簿
			//指定工作薄eg:
			//List<ExcelRow> dataList2 = ExcelHelper.parseExcelRowList(excelInput,0);//默认只读取sheetIndex=0的工作簿数据，第二个参数指定工作簿
			System.out.println("---------------------------excel转系统自带ExcelRow对象----------------------");
			System.out.println(JSON.toJSONString(dataList1));
			
			/***********返回一个自定义对象List<Person>类型数据，需要提前定义excel模板文件，如测试中的demo1Templete.xlsx*********************************/
			InputStream templeteInput = Test1.class.getResourceAsStream("demo1Templete.xlsx");
			List<ExcelRow> templeteList1 = ExcelHelper.parseExcelRowList(templeteInput);
			List<Person> objList = ExcelHelper.parseExcelToObject(dataList1, templeteList1, Person.class);
			System.out.println("---------------------------excel转对象----------------------");
			System.out.println(JSON.toJSONString(objList));
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
