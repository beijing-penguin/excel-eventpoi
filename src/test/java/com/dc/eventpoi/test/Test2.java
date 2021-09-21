package com.dc.eventpoi.test;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

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
            p.setName("");
            p.setNo("123");
            p.setRemark("备注");
            Files.write(Paths.get("./my_test_temp/file.xlsx"), ExcelHelper.exportExcel(output.toByteArray(), p));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
