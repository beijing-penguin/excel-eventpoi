package com.dc.eventpoi.test;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import com.dc.eventpoi.ExcelHelper;
import com.dc.eventpoi.core.PoiUtils;

public class 使用一行代码导出excel {
    public static void main(String[] args) throws Exception {
        List<Person> personList = new ArrayList<Person>();
        //构造导出时的数据
        for (int i = 0; i < 2000; i++) {
            
            Person p1 = new Person();
            p1.setNo("NO_"+i);
            p1.setAge(11);
            p1.setName("ssssss_"+i);
            p1.setRemark("测试测试啊remar_"+i);
            personList.add(p1);
        }
        
        //删除某些列
        byte[] newTempFile = PoiUtils.deleteCol(Test1.class.getResourceAsStream("demo1Templete.xlsx"), "${salary}");
        long t1 = System.currentTimeMillis();
        byte[] exportByteData = ExcelHelper.exportExcel(newTempFile, personList,0);
        System.out.println("cost="+(System.currentTimeMillis()-t1));
        //支持设置单元格样式噢！！！^_^
        //        byte[] exportByteData = ExcelHelper.exportExcel(Test1.class.getResourceAsStream("demo1Templete.xlsx"), personList,new CallBackCellStyle() {
        //            @Override
        //            public void callBack(CellStyle cellStyle) {
        //              cellStyle.setAlignment(HorizontalAlignment.CENTER); // 水平居中
        //              cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 上下居中
        //              cellStyle.setBorderTop(BorderStyle.THIN);
        //              cellStyle.setBorderBottom(BorderStyle.THIN);
        //              cellStyle.setBorderLeft(BorderStyle.THIN);
        //              cellStyle.setBorderRight(BorderStyle.THIN);
        //            }
        //        }, "${salary}");
        Files.write(Paths.get("./my_test_temp/测试导出指定对象并删除指定列.xlsx"), exportByteData);
    }
}
