package com.dc.eventpoi.test.temp;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import com.dc.eventpoi.core.PoiUtils;
import com.dc.eventpoi.core.entity.ListAndTableEntity;
import com.dc.eventpoi.test.Me;
import com.dc.eventpoi.test.temp.write.XlsxWriteStream;

public class TestWriteXlsx {

	public static void main(String[] args) throws Throwable  {
		InputStream tempInputStream = Me.class.getResourceAsStream("demo1Templete.xlsx");
		
		Student s1 = new Student();
		s1.setName(null);
		s1.setAge(3);
		
		Student s2 = new Student();
		s2.setName("李四");
		s2.setAge(4);
		
		ListAndTableEntity tt = new ListAndTableEntity();
		List<Student> stuList = new ArrayList<>();
		for (int i = 0; i < 1000; i++) {
			stuList.add(s1);
		}
		tt.setDataList(stuList);
		tt.setTable(s1);
		
		long t1 = System.currentTimeMillis();
		XlsxWriteStream ww = new XlsxWriteStream();
		ww.exportExcel(PoiUtils.inputStreamToByte(tempInputStream), tt);
		System.err.println("cost="+(System.currentTimeMillis()-t1));
	}
}
class Student{
	
	private String name;
	private Integer age;
	
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
}