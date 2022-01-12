package com.dc.eventpoi.test;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import com.dc.eventpoi.ExcelHelper;
import com.dc.eventpoi.core.PoiUtils;
import com.dc.eventpoi.core.entity.ListAndTableEntity;

public class 使用一行代码导出excel {
	public static void main(String[] args) throws Exception {
		String img_file_path = new File(Me.class.getResource("unnamed.jpg").getPath()).getAbsolutePath();
		byte[] imageData = Files.readAllBytes(Paths.get(img_file_path));
		//写到本地，并查看图片
		Files.write(Paths.get("./my_test_temp/unnamed_temp.jpg"), imageData);

		List<Person> personList = new ArrayList<Person>();
		//构造导出时的数据
		for (int i = 0; i < 20; i++) {

			Person p1 = new Person();
			p1.setNo("NO_"+i);
			p1.setAge(11);
			p1.setName("ssssss_"+i);
			p1.setRemark("测试测试啊remar_"+i);
			if(i==0) {
				p1.setHeadImage(imageData);
			}
			personList.add(p1);
		}

		//模拟一些需要删除某些列 的业务场景
		//byte[] newTempFile = PoiUtils.deleteCol(Me.class.getResourceAsStream("demo1Templete.xlsx"), "${salary}");
		
		byte[] tempData = PoiUtils.inputStreamToByte(Me.class.getResourceAsStream("demo1Templete.xlsx"));
		ListAndTableEntity dataEntity = ListAndTableEntity.build().setDataList(personList);
		
		long t1 = System.currentTimeMillis();
		byte[] exportByteData = ExcelHelper.exportExcel(tempData, dataEntity,null,null,null,null);
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
