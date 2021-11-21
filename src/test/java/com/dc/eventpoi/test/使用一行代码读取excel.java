package com.dc.eventpoi.test;

import java.util.List;

import com.alibaba.fastjson.JSON;
import com.dc.eventpoi.ExcelHelper;

public class 使用一行代码读取excel {
    public static void main(String[] args) throws Exception {
        List<Person> objList = ExcelHelper.parseExcelToObject(Me.class.getResourceAsStream("demo1Templete.xlsx"),Me.class.getResourceAsStream("demo1.xlsx"),  Person.class);
        System.err.println(JSON.toJSONString(objList,true));
    }
}
