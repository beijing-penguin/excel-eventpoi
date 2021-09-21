package com.dc.eventpoi.test;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import com.dc.eventpoi.ExcelHelper;

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
        List<Object>  excelDataList = new ArrayList<Object>();
        List<ProductInfo> productList = new ArrayList<ProductInfo>();
        ProductInfo p1 = new ProductInfo();
        p1.setNo("NO_1");
        p1.setName("ssssss111");
        String img_file_path = new File(Test1.class.getResource("unnamed.jpg").getPath()).getAbsolutePath();
        p1.setHeadImage(Files.readAllBytes(Paths.get(img_file_path)));
        
        ProductInfo p2 = new ProductInfo();
        p2.setNo("NO_2");
        p2.setName("ssssss222");
        
        ProductInfo p3 = new ProductInfo();
        p3.setNo("NO_3");
        p3.setName("ssssss333");
        
        productList.add(p1);
        productList.add(p2);
        productList.add(p3);
        
        excelDataList.add(productList);
        excelDataList.add(orderInfo);

        //第三个参数表示，导出时，删除那些列（按模板文件中的key删除，可不传）
        byte[] exportByteData = ExcelHelper.exportExcel(Test1.class.getResourceAsStream("订单_templete.xlsx"), excelDataList,0);
        Files.write(Paths.get("./my_test_temp/测试包含表格和列表数据的复杂导出.xlsx"), exportByteData);
    }
}
