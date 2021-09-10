package com.dc.eventpoi.test;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import com.dc.eventpoi.ExcelHelper;

public class 使用一行代码导出excel {
    public static void main(String[] args) throws Exception {
        //构造导出时的数据（带图片的数据）
        List<Person> personList = new ArrayList<Person>();
        Person p1 = new Person();
        p1.setNo("NO_1");
        p1.setAge(11);
        p1.setName("ssssss111");
        p1.setRemark("测试测试啊remark111");
        String img_file_path = new File(Test1.class.getResource("unnamed.jpg").getPath()).getAbsolutePath();
        p1.setHeadImage(Files.readAllBytes(Paths.get(img_file_path)));
        
        Person p2 = new Person();
        p2.setNo("NO_2");
        p2.setAge(22);
        p2.setName("ssssss222");
        p2.setRemark("测试测试啊remark2222");
        String img_file_path2 = new File(Test1.class.getResource("20200729144457_84047.jpg").getPath()).getAbsolutePath();
        p2.setHeadImage(Files.readAllBytes(Paths.get(img_file_path2)));
        
        Person p3 = new Person();
        p3.setNo("NO_3");
        p3.setAge(333);
        p3.setName("ssssss333");
        p3.setRemark("测试测试啊remark3333");
        p3.setHeadImage("文本数据测试333".getBytes());
        
        personList.add(p1);
        personList.add(p2);
        personList.add(p3);

        //第三个参数表示，导出时，删除那些列（按模板文件中的key删除，可不传）
        byte[] exportByteData = ExcelHelper.exportExcel(Test1.class.getResourceAsStream("demo1Templete.xlsx"), personList, "${salary}");
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
