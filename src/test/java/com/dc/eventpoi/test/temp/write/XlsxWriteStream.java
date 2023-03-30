/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package com.dc.eventpoi.test.temp.write;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dc.eventpoi.test.Me;
import com.dc.eventpoi.test.temp.read.CellReadCallBack;
import com.dc.eventpoi.test.temp.read.SheetReadCallBack;
import com.dc.eventpoi.test.temp.read.XlsxReadStream;


public class XlsxWriteStream {
	
			
	public void exportExcel(byte[] tempExcelBtye,Set<Object> listAndTableSet) throws Throwable {
		//XlsxReadStream xlsxRead = new XlsxReadStream();
		//LinkedHashMap<SheetReadCallBack,List<CellReadCallBack>> tempList = xlsxRead.setFileInputStream(new ByteArrayInputStream(tempExcelBtye)).doRead();
		
		XSSFWorkbook tempWorkbook = new XSSFWorkbook(new ByteArrayInputStream(tempExcelBtye));
		
		LinkedHashMap<Sheet,List<Cell>> tempList = new LinkedHashMap<>();
		
		
		int sheetEnd = tempWorkbook.getNumberOfSheets();
		for (int sheetIndex = 0; sheetIndex < sheetEnd; sheetIndex++) {
			Sheet sheet = tempWorkbook.getSheetAt(sheetIndex);
			int rowEnd = sheet.getPhysicalNumberOfRows();
			for (int rowIndex = 0; rowIndex < rowEnd; rowIndex++) {
				Row row = sheet.getRow(rowIndex);
			}
		}
		
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
	
	public int findEndSheetIndex(Workbook workbook) {
		return workbook.getNumberOfSheets();
	}
	public Row findRow(Workbook workbook,int sheetIndex,int rowIndex) {
		return workbook.getSheetAt(sheetIndex).getRow(rowIndex);
	}
	
}
