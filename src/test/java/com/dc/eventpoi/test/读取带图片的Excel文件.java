package com.dc.eventpoi.test;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

import com.alibaba.fastjson.JSON;
import com.dc.eventpoi.ExcelHelper;

public class 读取带图片的Excel文件 {
    public static void main(String[] args) throws Exception {
        List<Person> objList = ExcelHelper.parseExcelToObject(Test1.class.getResourceAsStream("demo1.xlsx"), Test1.class.getResourceAsStream("demo1Templete.xlsx"), Person.class,true);
        System.out.println(JSON.toJSONString(objList,true));
        
        System.err.println(new String(objList.get(2).getHeadImage()));
        //写到本地，并查看图片
        Files.write(Paths.get("./my_test_temp/my_image.jpg"), objList.get(0).getHeadImage());
    }
}
