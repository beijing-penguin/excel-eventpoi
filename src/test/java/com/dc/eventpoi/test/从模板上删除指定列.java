package com.dc.eventpoi.test;

import java.nio.file.Files;
import java.nio.file.Paths;

import com.dc.eventpoi.core.PoiUtils;

public class 从模板上删除指定列 {
    public static void main(String[] args) throws Exception {
        byte[] fileData = PoiUtils.deleteTemplateColumn(Me.class.getResourceAsStream("demo1Templete.xlsx"),"${salary}");
        Files.write(Paths.get("./my_test_temp/从模板上删除指定列.xlsx"), fileData);
    }
}
