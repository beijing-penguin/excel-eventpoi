package com.dc.eventpoi.test.temp;

import java.io.InputStream;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;

import com.dc.eventpoi.core.PoiUtils;
import com.dc.eventpoi.test.Me;
import com.dc.eventpoi.test.temp.write.XlsxWriteStream;

public class TestWriteXlsx {

	public static void main(String[] args) throws Throwable  {
		InputStream tempInputStream = Me.class.getResourceAsStream("demo1Templete.xlsx");
		
		Student s1 = new Student();
		s1.setName("张三");
		
		Set<Object> ss = new HashSet<>();
		ss.add(s1);
		
		XlsxWriteStream ww = new XlsxWriteStream();
		ww.exportExcel(PoiUtils.inputStreamToByte(tempInputStream), Arrays.asList(s1));
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