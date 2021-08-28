package com.dc.eventpoi.test;

import java.io.FileOutputStream;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.dc.eventpoi.ExcelHelper;

public class 使用列索引删除列测试 {
    public static void main(String[] args) throws Exception {
        InputStream excelInput = Test1.class.getResourceAsStream("demo1.xlsx");
        Workbook workbook = WorkbookFactory.create(excelInput);
        Sheet sheet = workbook.getSheetAt(0);
        
        ExcelHelper.deleteColumn(sheet, 1);
        
        FileOutputStream fileOut = new FileOutputStream("./my_test_temp/删除列.xlsx");
        workbook.write(fileOut);
        fileOut.flush();
    }
}
