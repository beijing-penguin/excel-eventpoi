package com.dc.eventpoi.test;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import com.dc.eventpoi.ExcelHelper;

public class 测试包含表格和列表数据的复杂导出 {
    public static void main(String[] args) throws Exception {
    	//列表或 表格的数据的集合
        List<Object>  excelDataList = new ArrayList<Object>();
        
    	//构造表格形式的数据
    	OrderInfo orderInfo = new OrderInfo();
    	orderInfo.setKehu("ddddcccc");
    	orderInfo.setOrderName("进口海鲜");
    	orderInfo.setTotalMoney("15.66");
    	orderInfo.setBuyer("张三");
    	orderInfo.setSaller("李四");
    	//条件表格数据
    	excelDataList.add(orderInfo);
    	
    	
        List<ProductInfo> productList = new ArrayList<ProductInfo>();
        //构造导出时的数据
        for (int i = 0; i < 20; i++) {
            
        	ProductInfo p1 = new ProductInfo();
            p1.setNo("NO_"+i);
            p1.setName("ssssss_"+i);
            if(i==0) {//测试用例，只导出第一行带图片的数据。
            	String img_file_path = new File(Test1.class.getResource("unnamed.jpg").getPath()).getAbsolutePath();
            	p1.setHeadImage(Files.readAllBytes(Paths.get(img_file_path)));
            }
            p1.setCaigouNum(i+10);
            productList.add(p1);
        }
        //条件列表数据
        excelDataList.add(productList);

        //第三个参数表示，导出时，删除那些列（按模板文件中的key删除，可不传）
        byte[] exportByteData = ExcelHelper.exportExcel(Test1.class.getResourceAsStream("订单_templete.xlsx"), excelDataList,0);
        Files.write(Paths.get("./my_test_temp/测试包含表格和列表数据的复杂导出.xlsx"), exportByteData);
    }
}
