package com.dc.eventpoi.test;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import com.dc.eventpoi.ExcelHelper;
import com.dc.eventpoi.core.CallBackCellStyle;

public class 测试导出指定对象并删除指定列 {
    public static void main(String[] args) throws Exception {
        List<Person> personList = new ArrayList<Person>();
        Person p1 = new Person();
        p1.setAge(11);
        p1.setName("ssssss111");
        p1.setRemark("测试测试啊remark111");
        
        Person p2 = new Person();
        p2.setAge(22);
        p2.setName("ssssss222");
        p2.setRemark("测试测试啊remark2222");
        personList.add(p1);
        personList.add(p2);
        

        //支持设置单元格样式噢！！！^_^
        byte[] exportByteData = ExcelHelper.exportExcel(Test1.class.getResourceAsStream("demo1Templete.xlsx"), personList,new CallBackCellStyle() {
            @Override
            public void callBack(CellStyle cellStyle) {
              cellStyle.setAlignment(HorizontalAlignment.CENTER); // 水平居中
              cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 上下居中
              cellStyle.setBorderTop(BorderStyle.THIN);
              cellStyle.setBorderBottom(BorderStyle.THIN);
              cellStyle.setBorderLeft(BorderStyle.THIN);
              cellStyle.setBorderRight(BorderStyle.THIN);
            }
        }, "${salary}");
        Files.write(Paths.get("./my_test_temp/测试导出指定对象并删除指定列.xlsx"), exportByteData);
    }
}
