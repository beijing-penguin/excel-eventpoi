package com.dc.eventpoi.test;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

import org.apache.commons.compress.utils.Lists;

import com.dc.eventpoi.ExcelHelper;

public class Test2 {
    public static void main(String[] args) {
        /***********返回List<ExcelRow>类型数据*********************************/
        InputStream excelInput = Test2.class.getResourceAsStream("客户收款通知书模板.xlsx");

        try {
            ByteArrayOutputStream output = new ByteArrayOutputStream();
            byte[] buffer = new byte[1024*4];
            int n = 0;
            while (-1 != (n = excelInput.read(buffer))) {
                output.write(buffer, 0, n);
            }
            
            Person p = new Person();
            p.setName("dc");
            p.setNo("123");
            List<Person> list = Lists.newArrayList();
            list.add(p);
            Files.write(Paths.get("C:\\Users\\Administrator\\Desktop\\file.xlsx"), ExcelHelper.exportTitleExcel(output.toByteArray(), p,0));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
