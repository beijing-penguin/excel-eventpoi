package com.dc.eventpoi.test.temp;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.dc.eventpoi.test.Me;
import com.dc.eventpoi.test.temp.read.CellReadCallBack;
import com.dc.eventpoi.test.temp.read.XlsxReadStream;

public class TestWriteXlsx {

	public static void main(String[] args) throws Throwable  {
		
		InputStream tempInputStream = Me.class.getResourceAsStream("demo1Templete.xlsx");
		System.err.println(tempInputStream);
		
		XlsxReadStream xlsxRead = new XlsxReadStream();
		List<CellReadCallBack> tempList = xlsxRead.setFileInputStream(tempInputStream).doRead().values().iterator().next();
		
		SXSSFWorkbook wb = new SXSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
		Sheet sh = wb.createSheet();
		for(int rownum = 0; rownum < 1000; rownum++){
			Row row = sh.createRow(rownum);
			for(int cellnum = 0; cellnum < 10; cellnum++){
				Cell cell = row.createCell(cellnum);
				String address = new CellReference(cell).formatAsString();
				cell.setCellValue(address);
			}
		}
		FileOutputStream out = new FileOutputStream("my_test_temp/sxssf.xlsx");
		wb.write(out);
		out.close();
		// dispose of temporary files backing this workbook on disk
		wb.dispose();
		wb.close();
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