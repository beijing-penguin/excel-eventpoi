package com.dc.eventpoi.test.temp;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.dc.eventpoi.core.PoiUtils;
import com.dc.eventpoi.test.Me;
import com.dc.eventpoi.test.temp.read.CellReadCallBack;
import com.dc.eventpoi.test.temp.read.XlsxReadStream;
import com.dc.eventpoi.test.temp.write.XlsxWriteStream;

public class TestWriteXlsx {

	public static void main(String[] args) throws Throwable  {
		InputStream tempInputStream = Me.class.getResourceAsStream("demo1Templete.xlsx");
		
		Student s1 = new Student();
		
		Set<Object> ss = new HashSet<>();
		ss.add(s1);
		
		XlsxWriteStream ww = new XlsxWriteStream();
		ww.exportExcel(PoiUtils.inputStreamToByte(tempInputStream), ss);
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