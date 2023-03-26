package com.dc.eventpoi.test.temp;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.DeferredSXSSFSheet;
import org.apache.poi.xssf.streaming.DeferredSXSSFWorkbook;

/**
 * https://poi.apache.org/components/spreadsheet/how-to.html#sxssf
 * @author beijing-penguin
 *
 */
public class 测试sxssf {
	public static void main(String[] args) throws Throwable {
		try (DeferredSXSSFWorkbook wb = new DeferredSXSSFWorkbook()) {
            DeferredSXSSFSheet sheet1 = wb.createSheet("new sheet");

            // cell styles should be created outside the row generator function
            CellStyle cellStyle = wb.createCellStyle();
            cellStyle.setAlignment(HorizontalAlignment.CENTER);

            sheet1.setRowGenerator((ssxSheet) -> {
                for (int i = 0; i < 10; i++) {
                    Row row = ssxSheet.createRow(i);
                    Cell cell = row.createCell(1);
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue("value " + i);
                }
            });

            try (FileOutputStream fileOut = new FileOutputStream("my_test_temp/DeferredGeneration.xlsx")) {
                wb.write(fileOut);
                //writeAvoidingTempFiles was added as an experimental change in POI 5.1.0
                //wb.writeAvoidingTempFiles(fileOut);
            } finally {
                //the dispose call is necessary to ensure temp files are removed
                wb.dispose();
            }
            System.out.println("wrote DeferredGeneration.xlsx");
        }
	}
}
