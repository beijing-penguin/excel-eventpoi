package com.dc.eventpoi.core;

import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import com.dc.eventpoi.ExcelHelper;
import com.dc.eventpoi.core.entity.ExcelCell;
import com.dc.eventpoi.core.entity.ExcelRow;
import com.dc.eventpoi.core.enums.FileType;

/**
 * POI工具
 * @author beijing-penguin
 *
 */
public class PoiUtils {
	/**
	 * 判断文件类型
	 *
	 * @param inp 数据流
	 * @return FileType
	 * @throws Exception IOException
	 * @author beijing-penguin
	 */
	public static FileType judgeFileType(InputStream inp) throws Exception {
		InputStream is = FileMagic.prepareToCheckMagic(inp);
		FileMagic fm = FileMagic.valueOf(is);

		switch (fm) {
		case OLE2:
			return FileType.XLS;
		case OOXML:
			return FileType.XLSX;
		default:
			throw new IOException("Your InputStream was neither an OLE2 stream, nor an OOXML stream");
		}
	}
	
	public static String getCellValue(Cell cell){
		String xssCellValue = null;
        if(cell.getCellType() == CellType.NUMERIC) {
        	xssCellValue = String.valueOf(cell.getNumericCellValue());
        }else {
        	xssCellValue = cell.getStringCellValue();
        }
        return xssCellValue;
	}
	
	public static Map<String, byte[]> getXlsxPictures(int sheetIndex, XSSFSheet sheet) throws Exception {
		Map<String, byte[]> map = new HashMap<String, byte[]>();
		List<POIXMLDocumentPart> list = sheet.getRelations();
		for (POIXMLDocumentPart part : list) {
			if (part instanceof XSSFDrawing) {
				XSSFDrawing drawing = (XSSFDrawing) part;
				List<XSSFShape> shapes = drawing.getShapes();
				for (XSSFShape shape : shapes) {
					XSSFPicture picture = (XSSFPicture) shape;
					XSSFClientAnchor anchor = picture.getPreferredSize();
					CTMarker marker = anchor.getFrom();
					String key = sheetIndex + "-" + marker.getRow() + "-" + marker.getCol();
					map.put(key, picture.getPictureData().getData());
				}
			}
		}
		return map;
	}

	public static Map<String, byte[]> getXlsPictures(int sheetIndex, HSSFSheet sheet) {
		Map<String, byte[]> map = new HashMap<String, byte[]>();
		List<HSSFShape> list = sheet.getDrawingPatriarch().getChildren();
		for (HSSFShape shape : list) {
			if (shape instanceof HSSFPicture) {
				HSSFPicture picture = (HSSFPicture) shape;
				HSSFClientAnchor cAnchor = picture.getClientAnchor();
				HSSFPictureData pdata = picture.getPictureData();
				String key = sheetIndex + "-" + cAnchor.getRow1() + "-" + cAnchor.getCol1(); // 行号-列号
				map.put(key, pdata.getData());
			}
		}
		return map;
	}
	
	public static String getImageType(byte[] b10) {
		byte b0 = b10[0];
		byte b1 = b10[1];
		byte b2 = b10[2];
		byte b3 = b10[3];
		byte b6 = b10[6];
		byte b7 = b10[7];
		byte b8 = b10[8];
		byte b9 = b10[9];
		if (b0 == (byte) 'G' && b1 == (byte) 'I' && b2 == (byte) 'F') {
			return "gif";
		} else if (b1 == (byte) 'P' && b2 == (byte) 'N' && b3 == (byte) 'G') {
			return "png";
		} else if (b6 == (byte) 'J' && b7 == (byte) 'F' && b8 == (byte) 'I' && b9 == (byte) 'F') {
			return "jpg";
		} else {
			return null;
		}
	}
	
	/**
	 * 常规自动日期格式识别
	 *
	 * @param str 时间字符串
	 * @return Date
	 * @author dc
	 */
	public static String getDateFormat(String str) {
		boolean year = false;
		Pattern pattern = Pattern.compile("^[-\\+]?[\\d]*$");
		if (pattern.matcher(str.substring(0, 4)).matches()) {
			year = true;
		}
		StringBuilder sb = new StringBuilder();
		int index = 0;
		if (!year) {
			if (str.contains("月") || str.contains("-") || str.contains("/")) {
				if (Character.isDigit(str.charAt(0))) {
					index = 1;
				}
			} else {
				index = 3;
			}
		}
		for (int i = 0; i < str.length(); i++) {
			char chr = str.charAt(i);
			if (Character.isDigit(chr)) {
				if (index == 0) {
					sb.append("y");
				} else if (index == 1) {
					sb.append("M");
				} else if (index == 2) {
					sb.append("d");
				} else if (index == 3) {
					sb.append("H");
				} else if (index == 4) {
					sb.append("m");
				} else if (index == 5) {
					sb.append("s");
				} else if (index == 6) {
					sb.append("S");
				}
			} else {
				if (i > 0) {
					char lastChar = str.charAt(i - 1);
					if (Character.isDigit(lastChar)) {
						index++;
					}
				}
				sb.append(chr);
			}
		}
		return sb.toString();
	}
	
	/**
	 * @param value     任意数据类型对象
	 * @param fieldType 转化后的类型
	 * @return Object
	 * @throws Exception IOException
	 * @author dc
	 */
	public static Object getValueByFieldType(Object value, Class<?> fieldType) throws Exception {
		if (value == null) {
			return null;
		}
		String v = String.valueOf(value);
		String type = fieldType.getSimpleName();
		if (type.equals("String")) {
			return v;
		} else if (v.trim().length() == 0) {
			return null;
		} else if (type.equals("Integer") || type.equals("int")) {
			return Integer.parseInt(v);
		} else if (type.equals("Long") || type.equals("long")) {
			return Long.parseLong(v);
		} else if (type.equals("Double") || type.equals("double")) {
			return Double.parseDouble(v);
		} else if (type.equals("Short") || type.equals("short")) {
			return Short.parseShort(v);
		} else if (type.equals("Float") || type.equals("float")) {
			return Float.parseFloat(v);
		} else if (type.equals("Byte") || type.equals("byte")) {
			return Byte.parseByte(v);
		} else if (type.equals("Byte[]") || type.equals("byte[]")) {
			return v.getBytes();
		} else if (type.equals("Boolean") || type.equals("boolean")) {
			return Boolean.parseBoolean(v);
		} else if (type.equals("BigDecimal")) {
			return new BigDecimal(v);
		} else if (type.equals("BigInteger")) {
			return new BigInteger(v);
		} else if (type.equals("Date")) {
			SimpleDateFormat sdf = new SimpleDateFormat(PoiUtils.getDateFormat(v));
			// 不允许底层java自动日期进行计算，直接抛出异常
			sdf.setLenient(false);
			Date date = sdf.parse(v);
			return date;
		}
		throw new Exception(type + " is unsupported");
	}
	
	/**
	 * 流转byte[]
	 *
	 * @param is is
	 * @return byte[]
	 * @throws Exception IOException
	 */
	public static byte[] inputStreamToByte(InputStream is) throws Exception {
		BufferedInputStream bis = new BufferedInputStream(is);
		byte[] buf = new byte[1024];
		int len = 0;
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		while ((len = bis.read(buf)) != -1) {
			bos.write(buf, 0, len);
		}
		bis.close();
		bos.flush();
		bos.close();
		return bos.toByteArray();
	}

	/**
	 * 获取cell值
	 *
	 * @param cellList    cell集合
	 * @param cellIndex   索引号
	 * @param returnClass 返回类型
	 * @param <T>         返回类型
	 * @return T
	 * @throws Exception IOException
	 */
	@SuppressWarnings("unchecked")
	public static <T> T getValueBy(List<ExcelCell> cellList, int cellIndex, Class<? extends T> returnClass) throws Exception {
		for (int i = 0; i < cellList.size(); i++) {
			ExcelCell cell = cellList.get(i);
			if (cell.getIndex() == cellIndex) {
				return (T) PoiUtils.getValueByFieldType(cell.getValue(), returnClass);
			}
		}
		return null;
	}

	/**
	 * 获取cell值
	 *
	 * @param cellList  cell集合
	 * @param cellIndex 索引号
	 * @return String
	 * @throws Exception IOException
	 */
	public static String getValueBy(List<ExcelCell> cellList, int cellIndex) throws Exception {
		return getValueBy(cellList, cellIndex, String.class);
	}

	/**
	 * 获取值
	 *
	 * @param rowList     行集合
	 * @param rowIndex    行下标
	 * @param cellIndex   列下标
	 * @param returnClass 返回值类型
	 * @param <T>         返回类型
	 * @return T
	 * @throws Exception IOException
	 */
	@SuppressWarnings("unchecked")
	public static <T> T getValueBy(List<ExcelRow> rowList, int rowIndex, int cellIndex, Class<? extends T> returnClass) throws Exception {
		for (int i = 0; i < rowList.size(); i++) {
			ExcelRow row = rowList.get(i);
			if (row.getRowIndex() > rowIndex) {
				break;
			}
			if (row.getRowIndex() == rowIndex) {
				List<ExcelCell> cellList = row.getCellList();
				for (int j = 0; j < cellList.size(); j++) {
					ExcelCell cell = cellList.get(j);
					if (cell.getIndex() == cellIndex) {
						return (T) PoiUtils.getValueByFieldType(cell.getValue(), returnClass);
					}
				}
			}
		}
		return null;
	}

	/**
	 * 获取值
	 *
	 * @param rowList   行集合
	 * @param rowIndex  行下标
	 * @param cellIndex 列下标
	 * @return String
	 * @throws Exception Exception
	 */
	public static String getValueBy(List<ExcelRow> rowList, int rowIndex, int cellIndex) throws Exception {
		return getValueBy(rowList, rowIndex, cellIndex, String.class);
	}

	public static void deleteColumn(Sheet sheet, int columnToDeleteIndex) {
		for (int rId = 0; rId <= sheet.getLastRowNum(); rId++) {
			Row row = sheet.getRow(rId);
			for (int cID = columnToDeleteIndex; cID < row.getLastCellNum(); cID++) {
				Cell cOld = row.getCell(cID);
				if (cOld != null) {
					row.removeCell(cOld);
				}
				Cell cNext = row.getCell(cID + 1);
				if (cNext != null) {
					Cell cNew = row.createCell(cID, cNext.getCellType());
					cloneCell(cNew, cNext);
					// Set the column width only on the first row.
					// Other wise the second row will overwrite the original column width set
					// previously.
					if (rId == 0) {
						sheet.setColumnWidth(cID, sheet.getColumnWidth(cID + 1));
					}
				}
			}
		}
	}

	public static void cloneCell(Cell cNew, Cell cOld) {
		cNew.setCellComment(cOld.getCellComment());
		cNew.setCellStyle(cOld.getCellStyle());

		if (CellType.BOOLEAN == cNew.getCellType()) {
			cNew.setCellValue(cOld.getBooleanCellValue());
		} else if (CellType.NUMERIC == cNew.getCellType()) {
			cNew.setCellValue(cOld.getNumericCellValue());
		} else if (CellType.STRING == cNew.getCellType()) {
			cNew.setCellValue(cOld.getStringCellValue());
		} else if (CellType.ERROR == cNew.getCellType()) {
			cNew.setCellValue(cOld.getErrorCellValue());
		} else if (CellType.FORMULA == cNew.getCellType()) {
			cNew.setCellValue(cOld.getCellFormula());
		}
	}

	public static byte[] deleteTemplateColumn(InputStream excelFileInput, String... key) throws Exception {
		ByteArrayOutputStream output = new ByteArrayOutputStream();
		byte[] buff = new byte[1024 * 4];
		int n = 0;
		while (-1 != (n = excelFileInput.read(buff))) {
			output.write(buff, 0, n);
		}
		return deleteTemplateColumn(output.toByteArray(), 0, key);
	}

	public static byte[] deleteTemplateColumn(InputStream excelFileInput, int sheetIndex, String... keys) throws Exception {
		ByteArrayOutputStream output = new ByteArrayOutputStream();
		byte[] buff = new byte[1024 * 4];
		int n = 0;
		while (-1 != (n = excelFileInput.read(buff))) {
			output.write(buff, 0, n);
		}
		return deleteTemplateColumn(output.toByteArray(), sheetIndex, keys);
	}

	public static byte[] deleteTemplateColumn(byte[] templateFile, String... key) throws Exception {
		return deleteTemplateColumn(templateFile, 0, key);
	}

	public static byte[] deleteTemplateColumn(byte[] templateFile, int sheetIndex, String... keys) throws Exception {
		Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(templateFile));
		Sheet sheet = workbook.getSheetAt(sheetIndex);

		List<ExcelRow> rowList = ExcelHelper.parseExcelRowList(new ByteArrayInputStream(templateFile));
		for (String kk : keys) {
			for (ExcelRow row : rowList) {
				List<ExcelCell> cellList = row.getCellList();
				for (ExcelCell cell : cellList) {
					if (cell.getValue().equals(kk) && sheetIndex == row.getSheetIndex()) {
						deleteColumn(sheet, cell.getIndex());
					}
				}
			}
		}

		ByteArrayOutputStream byteOut = new ByteArrayOutputStream();
		workbook.write(byteOut);
		byteOut.flush();
		byteOut.close();
		return byteOut.toByteArray();
	}
	
	/**
	 * 删除指定列
	 * 
	 * @param templeteStream 原始模板
	 * @param delCellKey     删除的列key，如${name}
	 * @return byte[]
	 * @throws Exception Exception
	 */
	public static byte[] deleteCol(InputStream templeteStream, String... delCellKey) throws Exception {
		ByteArrayOutputStream templeteOutput = new ByteArrayOutputStream();
		byte[] buff = new byte[1024 * 4];
		int n = 0;
		while (-1 != (n = templeteStream.read(buff))) {
			templeteOutput.write(buff, 0, n);
		}
		byte[] newTemPleteExcel = deleteTemplateColumn(templeteOutput.toByteArray(), delCellKey);
		return newTemPleteExcel;
	}
	
	/**
	 * 删除模板中的固定格式
	 *
	 * @param inputSrc   源模板文件
	 * @param sheetIndex 工作簿索引下标
	 * @return ByteArrayOutputStream
	 * @throws Exception IOException
	 */
	public static ByteArrayOutputStream deleteTempleteFormat(InputStream inputSrc, int sheetIndex) throws Exception {
		Workbook workbook = WorkbookFactory.create(inputSrc);
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		int totalRow = sheet.getPhysicalNumberOfRows();
		for (int i = sheet.getFirstRowNum(); i < totalRow; i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				for (int j = row.getFirstCellNum(), totalCell = row.getPhysicalNumberOfCells(); j < totalCell; j++) {
					Cell cell = row.getCell(j);
					if (cell != null) {
						// cell.setCellType(CellType.STRING);
						String value = cell.getStringCellValue();
						if (value != null && value.startsWith("${")) {
							sheet.removeRow(row);
							sheet.shiftRows(i + 1, i + 1 + 1, -1);
							break;
						}
					}
				}
			}
		}
		ByteArrayOutputStream os = new ByteArrayOutputStream();
		workbook.write(os);
		os.flush();
		return os;
	}
//
//
//	/**
//	 * 复制一个单元格样式到目的单元格样式
//	 * @param fromStyle
//	 * @param toStyle
//	 */
//	public static void copyCellStyle(CellStyle fromStyle, CellStyle toStyle) {
//		toStyle.setAlignment(fromStyle.getAlignment());
//		//边框和边框颜色
//		toStyle.setBorderBottom(fromStyle.getBorderBottom());
//		toStyle.setBorderLeft(fromStyle.getBorderLeft());
//		toStyle.setBorderRight(fromStyle.getBorderRight());
//		toStyle.setBorderTop(fromStyle.getBorderTop());
//		toStyle.setTopBorderColor(fromStyle.getTopBorderColor());
//		toStyle.setBottomBorderColor(fromStyle.getBottomBorderColor());
//		toStyle.setRightBorderColor(fromStyle.getRightBorderColor());
//		toStyle.setLeftBorderColor(fromStyle.getLeftBorderColor());
//
//		//背景和前景
//		toStyle.setFillBackgroundColor(fromStyle.getFillBackgroundColor());
//		toStyle.setFillForegroundColor(fromStyle.getFillForegroundColor());
//
//		toStyle.setDataFormat(fromStyle.getDataFormat());
//		toStyle.setFillPattern(fromStyle.getFillPattern());
//		//		toStyle.setFont(fromStyle.getFont(null));
//		toStyle.setHidden(fromStyle.getHidden());
//		toStyle.setIndention(fromStyle.getIndention());//首行缩进
//		toStyle.setLocked(fromStyle.getLocked());
//		toStyle.setRotation(fromStyle.getRotation());//旋转
//		toStyle.setVerticalAlignment(fromStyle.getVerticalAlignment());
//		toStyle.setWrapText(fromStyle.getWrapText());
//
//	}
//	/**
//	 * Sheet复制
//	 * @param fromSheet
//	 * @param toSheet
//	 * @param copyValueFlag
//	 */
//	public static void copySheet(SXSSFWorkbook wb,XSSFSheet fromSheet, SXSSFSheet toSheet, boolean copyValueFlag) {
//		//合并区域处理
//		mergerRegion(fromSheet, toSheet);
//		for (Iterator<Row> rowIt = fromSheet.rowIterator(); rowIt.hasNext();) {
//			XSSFRow tmpRow = (XSSFRow) rowIt.next();
//			SXSSFRow newRow = toSheet.createRow(tmpRow.getRowNum());
//			//行复制
//			copyRow(wb,tmpRow,newRow,copyValueFlag);
//		}
//	}
//	/**
//	 * 行复制功能
//	 * @param fromRow
//	 * @param toRow
//	 */
//	public static void copyRow(SXSSFWorkbook wb,XSSFRow fromRow,SXSSFRow toRow,boolean copyValueFlag){
//		for (Iterator<Cell> cellIt = fromRow.cellIterator(); cellIt.hasNext();) {
//			XSSFCell tmpCell = (XSSFCell) cellIt.next();
//			SXSSFCell newCell = toRow.createCell(tmpCell.getColumnIndex());
//			copyCell(wb,tmpCell, newCell, copyValueFlag);
//		}
//	}
//	/**
//	 * 复制原有sheet的合并单元格到新创建的sheet
//	 * 
//	 * @param sheetCreat 新创建sheet
//	 * @param sheet      原有的sheet
//	 */
//	public static void mergerRegion(XSSFSheet fromSheet, SXSSFSheet toSheet) {
//		int sheetMergerCount = fromSheet.getNumMergedRegions();
//		for (int i = 0; i < sheetMergerCount; i++) {
//			CellRangeAddress mergedRegionAt = fromSheet.getMergedRegion(i);
//			toSheet.addMergedRegion(mergedRegionAt);
//		}
//	}
//	/**
//	 * 复制单元格
//	 * 
//	 * @param srcCell
//	 * @param distCell
//	 * @param copyValueFlag
//	 *            true则连同cell的内容一起复制
//	 */
//	public static void copyCell(SXSSFWorkbook wb,XSSFCell srcCell, SXSSFCell distCell, boolean copyValueFlag) {
//		
//		CellStyle newstyle=wb.createCellStyle();
//		copyCellStyle(srcCell.getCellStyle(), newstyle);
//		//distCell.setEncoding(srcCell.get);
//		//样式
//		distCell.setCellStyle(newstyle);
//		//评论
//		if (srcCell.getCellComment() != null) {
//			distCell.setCellComment(srcCell.getCellComment());
//		}
//		// 不同数据类型处理
//		CellType srcCellType = srcCell.getCellType();
//		distCell.setCellType(srcCellType);
//		if (copyValueFlag) {
//			if (srcCellType == CellType.NUMERIC) {
//				if (DateUtil.isCellDateFormatted(srcCell)) {
//					distCell.setCellValue(srcCell.getDateCellValue());
//				} else {
//					distCell.setCellValue(srcCell.getNumericCellValue());
//				}
//			} else if (srcCellType == CellType.STRING) {
//				distCell.setCellValue(srcCell.getRichStringCellValue());
//			} else if (srcCellType == CellType.BLANK) {
//				// nothing21
//			} else if (srcCellType == CellType.BOOLEAN) {
//				distCell.setCellValue(srcCell.getBooleanCellValue());
//			} else if (srcCellType == CellType.ERROR) {
//				distCell.setCellErrorValue(srcCell.getErrorCellValue());
//			} else if (srcCellType == CellType.FORMULA) {
//				distCell.setCellFormula(srcCell.getCellFormula());
//			} else { // nothing29
//			}
//		}
//	}
}
