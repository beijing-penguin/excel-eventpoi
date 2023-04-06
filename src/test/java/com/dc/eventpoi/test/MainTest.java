package com.dc.eventpoi.test;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import com.alibaba.fastjson.JSON;
import com.dc.eventpoi.ExcelHelper;
import com.dc.eventpoi.core.PoiUtils;
import com.dc.eventpoi.core.entity.ListAndTableEntity;

public class MainTest {
	public static void main(String[] args) throws Throwable {
		String data = "{\"orderRebateTimeSummary\":\"2021-11-29\",\"grossProfitTotalSummary\":\"10259\",\"salesmanCommissionTotalSummary\":\"6642\",\"technicalServiceMoneyTotalSummary\":\"609\",\"transactionServiceMoneyTotalSummary\":\"26\",\"netProfitTotalSummary\":\"2982\"}";
		ShopkeeperBillTotalSummary sumData = JSON.parseObject(data,ShopkeeperBillTotalSummary.class);
		List<ShopkeeperBillTotalSummary> ll = new ArrayList<ShopkeeperBillTotalSummary>();
		ll.add(sumData);
		
		ListAndTableEntity dataEntity = ListAndTableEntity.build().setDataList(ll);
		
		byte[] tempData = PoiUtils.inputStreamToByte(MainTest.class.getResourceAsStream("exportPcBillStaticList_templete.xlsx"));
		byte[] exportByteData = ExcelHelper.exportExcel(tempData, dataEntity, null, null, null, null);
		Files.write(Paths.get("./my_test_temp/exportPcBillStaticList_templete导出.xlsx"), exportByteData);
	}
}
