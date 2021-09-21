package com.dc.eventpoi.test;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.dc.eventpoi.ExcelHelper;

public class Test3 {
    public static void main(String[] args) {
        /***********返回List<ExcelRow>类型数据*********************************/
        InputStream excelInput = Test3.class.getResourceAsStream("客户收款通知书模板.xlsx");

        try {
            long t1 = System.currentTimeMillis();
            SXSSFWorkbook workbook = new SXSSFWorkbook();
            workbook.createSheet("aaa");
            SXSSFSheet aaa = workbook.getSheetAt(0);
            for (int i=0;i<1000;i++){
                aaa.createRow(i);
                aaa.getRow(i).createCell(0).setCellValue("aaaaaaaaaaaaaaaaaaaaaaa"+i);
                aaa.getRow(i).createCell(1).setCellValue("aaaaaaaaaaaaaaaaaaaaaaa");
                aaa.getRow(i).createCell(2).setCellValue("aaaaaaaaaaaaaaaaaaaaaaa");
                aaa.getRow(i).createCell(3).setCellValue("aaaaaaaaaaaaaaaaaaaaaaa");
                aaa.getRow(i).createCell(4).setCellValue("aaaaaaaaaaaaaaaaaaaaaaa");
                aaa.getRow(i).createCell(0).setCellValue("aaaaaaaaaaaaaaaaaaaaaaa"+i);
            }
            OutputStream outputStream = null;
            // 打开目的输入流，不存在则会创建
            outputStream = new FileOutputStream("./my_test_temp/test3.xlsx");
            workbook.write(outputStream);
            outputStream.flush();
            outputStream.close();
            workbook.close();
            long t2 = System.currentTimeMillis();
            System.out.println("SXSSFWorkbook : 100w条数据写入Excel 消耗时间："+ (t2-t1));
            
        } catch (Exception e) {
            e.printStackTrace();
        }
	
    }
}
