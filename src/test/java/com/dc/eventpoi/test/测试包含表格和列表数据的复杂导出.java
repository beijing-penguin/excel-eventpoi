package com.dc.eventpoi.test;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import com.dc.eventpoi.ExcelHelper;
import com.dc.eventpoi.core.PoiUtils;
import com.dc.eventpoi.core.entity.ListAndTableEntity;

public class 测试包含表格和列表数据的复杂导出 {
    public static void main(String[] args) throws Exception {
    	//构造表格形式的数据
    	OrderInfo orderInfo = new OrderInfo();
    	orderInfo.setKehu("ddddcccc");
    	orderInfo.setOrderName("进口海鲜");
    	orderInfo.setTotalMoney("15.66");
    	orderInfo.setBuyer("张三");
    	orderInfo.setSaller("李四");
    	
    	//构造列表形式的数据
        List<ProductInfo> productList = new ArrayList<ProductInfo>();
        //构造导出时的数据
        for (int i = 0; i < 10; i++) {
            
        	ProductInfo p1 = new ProductInfo();
            p1.setNo("NO_"+i);
            p1.setName("ssssss_"+i);
            if(i==0) {//测试用例，只导出第一行带图片的数据。
            	String img_file_path = new File(Me.class.getResource("unnamed.jpg").getPath()).getAbsolutePath();
            	p1.setHeadImage(Files.readAllBytes(Paths.get(img_file_path)));
            }
            p1.setCaigouNum(i+10);
            productList.add(p1);
        }

        ListAndTableEntity dataEntity = ListAndTableEntity.build().setDataList(productList).setTable(orderInfo);
        byte[] tempData = PoiUtils.inputStreamToByte(Me.class.getResourceAsStream("订单_templete.xlsx"));
        long t1 = System.currentTimeMillis();
        byte[] exportByteData = ExcelHelper.exportExcel(tempData,dataEntity,null,null,null,null);
        System.out.println("导出成功，耗时="+(System.currentTimeMillis()-t1)+"毫秒");
        Files.write(Paths.get("./my_test_temp/测试包含表格和列表数据的复杂导出.xlsx"), exportByteData);
    }
}
