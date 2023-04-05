package com.dc.eventpoi.test.temp;

import java.io.File;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import com.dc.eventpoi.core.PoiUtils;
import com.dc.eventpoi.core.entity.ListAndTableEntity;
import com.dc.eventpoi.test.Me;
import com.dc.eventpoi.test.temp.write.XlsxWriteStream;

public class TestWriteXlsx {

	public static void main(String[] args) throws Throwable  {
		InputStream tempInputStream = Me.class.getResourceAsStream("demo1Templete.xlsx");
		
		String img_file_path = new File(Me.class.getResource("unnamed.jpg").getPath()).getAbsolutePath();
		byte[] img_byte = Files.readAllBytes(Paths.get(img_file_path));
		
		ListAndTableEntity tt = new ListAndTableEntity();
		List<Student> stuList = new ArrayList<>();
		for (int i = 0; i < 10; i++) {
			
			Student s1 = new Student();
			s1.setName("张三"+i);
			s1.setAge(3);
			s1.setHeadImage(img_byte);
			
			
			Field[] v_obj_field_arr = s1.getClass().getDeclaredFields();
			for (Field field : v_obj_field_arr) {
				System.err.println(field.getType().getTypeName());
			}
			stuList.add(s1);
		}
		tt.setDataList(stuList);
		
		Student s2 = new Student();
		s2.setName("李四");
		s2.setAge(3);
		tt.setTable(s2);
		
		long t1 = System.currentTimeMillis();
		XlsxWriteStream ww = new XlsxWriteStream();
		ww.setAutoClearPlaceholder(true);
		ww.exportExcel(PoiUtils.inputStreamToByte(tempInputStream), tt);
		System.err.println("cost="+(System.currentTimeMillis()-t1));
	}
}
class Student{
	
	private String name;
	private Integer age;
	private byte[] headImage;
	
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public Integer getAge() {
		return age;
	}
	public void setAge(Integer age) {
		this.age = age;
	}
	public byte[] getHeadImage() {
		return headImage;
	}
	public void setHeadImage(byte[] headImage) {
		this.headImage = headImage;
	}
}