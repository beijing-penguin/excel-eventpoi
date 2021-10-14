package com.dc.eventpoi.test;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFDrawing;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;

public class 图片导出测试 {
	public static void main(String[] args) throws Throwable {
		SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook();
		SXSSFSheet sxssSheet = sxssfWorkbook.createSheet("aaaaa");
		SXSSFDrawing patriarch = (SXSSFDrawing) sxssSheet.createDrawingPatriarch();
		
		SXSSFRow sxssrow = sxssSheet.createRow(0);
		
		SXSSFCell _sxssCell = sxssrow.createCell(0);
		
		String img_file_path = new File(Test1.class.getResource("img.jpg").getPath()).getAbsolutePath();
		byte[] value = Files.readAllBytes(Paths.get(img_file_path));
		
		XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, _sxssCell.getColumnIndex(), sxssrow.getRowNum(), _sxssCell.getColumnIndex() + 1, sxssrow.getRowNum() + 1);
        int picIndex = sxssfWorkbook.addPicture((byte[]) value, HSSFWorkbook.PICTURE_TYPE_JPEG);
        patriarch.createPicture(anchor, picIndex);
//        
//		SXSSFDrawing drawing = sxssSheet.createDrawingPatriarch();
//        if (drawing == null) {
//            drawing = sxssSheet.createDrawingPatriarch();
//        }
//        
//        CreationHelper helper = sxssSheet.getWorkbook().getCreationHelper();
//        ClientAnchor anchor = helper.createClientAnchor();
//        anchor.setDx1(0);
//        anchor.setDx2(0);
//        anchor.setDy1(0);
//        anchor.setDy2(0);
//        anchor.setCol1(_sxssCell.getColumnIndex());
//        anchor.setCol2(_sxssCell.getColumnIndex() + 1);
//        anchor.setRow1(_sxssCell.getRowIndex());
//        anchor.setRow2(_sxssCell.getRowIndex() + 1);
//        
//        int picIndex = sxssSheet.getWorkbook().addPicture(value, HSSFWorkbook.PICTURE_TYPE_JPEG);
//        drawing.createPicture(anchor, picIndex);
        
        ByteArrayOutputStream byteStream = new ByteArrayOutputStream();
        sxssfWorkbook.write(byteStream);
        byteStream.flush();
        byteStream.close();
        sxssfWorkbook.close();
        
        
        Files.write(Paths.get("./my_test_temp/图片导出测试.xlsx"), byteStream.toByteArray());
	}
}
