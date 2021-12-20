package com.dc.eventpoi.test;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import com.alibaba.fastjson.JSON;
import com.dc.eventpoi.ExcelHelper;

public class MainTest {
	public static void main(String[] args) throws Exception {
		String data = "{\"orderRebateTimeSummary\":\"2021-11-29\",\"grossProfitTotalSummary\":\"10259\",\"salesmanCommissionTotalSummary\":\"6642\",\"technicalServiceMoneyTotalSummary\":\"609\",\"transactionServiceMoneyTotalSummary\":\"26\",\"netProfitTotalSummary\":\"2982\"}";
		ShopkeeperBillTotalSummary sumData = JSON.parseObject(data,ShopkeeperBillTotalSummary.class);
		List<ShopkeeperBillTotalSummary> ll = new ArrayList<ShopkeeperBillTotalSummary>();
		ll.add(sumData);
		byte[] exportByteData = ExcelHelper.exportExcel(MainTest.class.getResourceAsStream("exportPcBillStaticList_templete.xlsx"), ll);
		Files.write(Paths.get("./my_test_temp/测试导出指定对象并删除指定列.xlsx"), exportByteData);
	}
}
