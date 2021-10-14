package com.dc.eventpoi.core;

//import java.io.FileOutputStream;
//import java.io.InputStream;
//import java.io.OutputStream;
//
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.CellStyle;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.WorkbookFactory;
//import org.apache.poi.ss.util.CellRangeAddress;
//import org.apache.poi.xssf.streaming.SXSSFCell;
//import org.apache.poi.xssf.streaming.SXSSFRow;
//import org.apache.poi.xssf.streaming.SXSSFSheet;
//import org.apache.poi.xssf.streaming.SXSSFWorkbook;
//import org.apache.poi.xssf.usermodel.XSSFCell;
//import org.apache.poi.xssf.usermodel.XSSFRow;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//import com.dc.eventpoi.core.enums.FileType;

/**
 * 废弃暂时不用
 * @author DC
 *
 */
@Deprecated
public class AbstractWriteWorkbook {
//
//	private XSSFWorkbook xssfWorkbook;
//	private SXSSFWorkbook sxssfWorkbook;
//	private HSSFWorkbook hssfWorkbook;
//
//	private FileType fileType;
//
//	public AbstractWriteWorkbook(InputStream templateStream) throws Exception {
//		fileType = PoiUtils.judgeFileType(templateStream);
//		if (fileType == FileType.XLSX) {
//			xssfWorkbook = new XSSFWorkbook(templateStream);
//			sxssfWorkbook = new SXSSFWorkbook(1000);
//			
//			for (int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
//				SXSSFSheet sxssSheet = sxssfWorkbook.createSheet(xssfWorkbook.getSheetName(i));
//				XSSFSheet xsssheet = xssfWorkbook.getSheetAt(i);
//				//PoiUtils.copySheet(sxssfWorkbook, xsssheet, sxssSheet, true);
//				int sheetMergerCount = xsssheet.getNumMergedRegions();
//				for (int ii = 0; ii < sheetMergerCount;ii++) {
//					CellRangeAddress mergedRegionAt = xsssheet.getMergedRegion(ii);
//					sxssSheet.addMergedRegion(mergedRegionAt);
//				}
//				int rowNum = xsssheet.getPhysicalNumberOfRows();
//				for (int j = 0; j < rowNum; j++) {
//					SXSSFRow sxssrow = sxssSheet.createRow(j);
//					XSSFRow xssrow = xsssheet.getRow(j);
//					sxssrow.setHeight(xssrow.getHeight());
//					int xssCellNum = xssrow.getPhysicalNumberOfCells();
//					for (int k = 0; k < xssCellNum; k++) {
//						XSSFCell xssCell = xssrow.getCell(k);
//						sxssSheet.setColumnWidth(k, xsssheet.getColumnWidth(k));
//						if(xssCell == null) {
////							SXSSFCell sxssCell = sxssrow.createCell(k);
////							sxssCell.setCellValue("");
//						}else {
//							CellStyle sxssStyle = sxssfWorkbook.createCellStyle();
//							sxssStyle.cloneStyleFrom(xssCell.getCellStyle());
//							SXSSFCell sxssCell = sxssrow.createCell(k,xssCell.getCellType());
//							sxssCell.setCellStyle(sxssStyle);
//							sxssCell.setCellValue(xssCell.getStringCellValue());
//						}
//					}
//				}
//			}
////			OutputStream outputStream = null;
////			// 打开目的输入流，不存在则会创建
////			outputStream = new FileOutputStream("./my_test_temp/test3333.xlsx");
////			sxssfWorkbook.write(outputStream);
////			outputStream.flush();
////			outputStream.close();
////			sxssfWorkbook.close();
//		} else {
//			hssfWorkbook = (HSSFWorkbook) WorkbookFactory.create(templateStream);
//		}
//	}
//
//	public Sheet getSheetAt(int i) throws Exception {
//		if (fileType == FileType.XLSX) {
//			SXSSFSheet sheet = sxssfWorkbook.getSheetAt(i);
//			return sheet;
//		}else {
//			return hssfWorkbook.getSheetAt(i);
//		}
//	}
//
//	public int addPicture(byte[] value, int pictureTypeJpeg) {
//		if (fileType == FileType.XLSX) {
//			return sxssfWorkbook.addPicture((byte[]) value, HSSFWorkbook.PICTURE_TYPE_JPEG);
//		}else {
//			return hssfWorkbook.addPicture((byte[]) value, HSSFWorkbook.PICTURE_TYPE_JPEG);
//		}
//	}
//
//	public void close() throws Exception {
//		if (fileType == FileType.XLSX) {
//			xssfWorkbook.close();
//			xssfWorkbook = null;
//
//			sxssfWorkbook.dispose();
//			sxssfWorkbook.close();
//		}else {
//			hssfWorkbook.close();
//		}
//	}
//
//	public void write(OutputStream byteStream) throws Exception {
//		if (fileType == FileType.XLSX) {
//			sxssfWorkbook.write(byteStream);
//		}else {
//			hssfWorkbook.write(byteStream);
//		}
//	}
//
//	public void dispose() {
//		if (fileType == FileType.XLSX) {
//			sxssfWorkbook.dispose();
//		}
//	}

}
