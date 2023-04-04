package com.dc.eventpoi.test.temp.write;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFDrawing;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import com.dc.eventpoi.core.PoiUtils;
import com.dc.eventpoi.core.entity.ListAndTableEntity;
import com.dc.eventpoi.test.temp.read.CellReadCallBack;
import com.dc.eventpoi.test.temp.read.SheetReadCallBack;
import com.dc.eventpoi.test.temp.read.XlsxReadStream;
import com.googlecode.aviator.AviatorEvaluator;

/**
 * https://poi.apache.org/components/spreadsheet/how-to.html#sxssf
 * @author beijing-penguin
 *
 */
public class XlsxWriteStream {

	/**
	 * 默认前缀占位符
	 */
	public String defaultPlaceholderPrefix = "${";

	/**
	 * 默认后缀占位符
	 */
	public String defaultPlaceholderSuffix = "}";


	public void exportExcel(byte[] tempExcelBtye,ListAndTableEntity listAndTableEntity) throws Throwable {

		SXSSFWorkbook export_workbook = new SXSSFWorkbook(-1);//关闭自动刷到磁盘，按行数计算
		export_workbook.setCompressTempFiles(true);//临时文件将被压缩

		XSSFWorkbook tempWorkbook = new XSSFWorkbook(new ByteArrayInputStream(tempExcelBtye));

		this.copyTempWorkbookToSxssfWorkbook(tempWorkbook,export_workbook);

		//开始写入数据
		int export_sheetEnd = tempWorkbook.getNumberOfSheets();
		for (int export_sheetIndex = 0; export_sheetIndex < export_sheetEnd; export_sheetIndex++) {
			XlsxReadStream readTemp = new XlsxReadStream();
			readTemp.setFileInputStream(new ByteArrayInputStream(tempExcelBtye));
			readTemp.setReadSheetIndex(export_sheetIndex);
			List<CellReadCallBack> tempContentCollection = readTemp.doRead().values().iterator().next();

			SXSSFSheet export_sheet = export_workbook.getSheetAt(export_sheetIndex);

			int export_rowEnd = export_sheet.getPhysicalNumberOfRows();
			for (int export_rowIndex = 0; export_rowIndex < export_rowEnd; export_rowIndex++) {
				Row export_row = export_sheet.getRow(export_rowIndex);

				int export_cellEnd = export_row.getPhysicalNumberOfCells();
				int export_cell_last_num = export_row.getLastCellNum();
				export_cellEnd = export_cellEnd>export_cell_last_num?export_cellEnd:export_cell_last_num;
				
				for (int export_cellIndex = 0; export_cellIndex < export_cellEnd; export_cellIndex++) {

					Cell export_cell = export_row.getCell(export_cellIndex);
					if(export_cell != null) {
						String export_cell_value = PoiUtils.getCellValue(export_cell);
						String newCellValue = resolveCellValue(export_cell_value, listAndTableEntity,tempContentCollection,export_rowIndex);
						System.err.println("newCellValue="+newCellValue);
						export_cell.setCellValue(newCellValue);
					}
				}

				//是否复制行
				List<CellReadCallBack> copyRowList = this.findCopyRow(tempContentCollection,listAndTableEntity,export_rowIndex);
				if(copyRowList != null) {
					export_rowEnd = export_rowEnd +1;
					SXSSFRow copy_row = export_sheet.createRow(export_rowIndex+1);
					copy_row.setHeight(export_sheet.getRow(export_rowIndex-1).getHeight());
					
					for(CellReadCallBack cell : copyRowList) {
						int rowCopyIndex = 0;
						if(cell.getRowIndex() - export_rowIndex == 0) {
							rowCopyIndex = cell.getRowIndex();
						}else {
							rowCopyIndex = export_rowIndex - 1;
						}
						System.err.println(rowCopyIndex);
						SXSSFCell copyCelll = copy_row.createCell(cell.getCellIndex(), export_sheet.getRow(rowCopyIndex).getCell(cell.getCellIndex()).getCellType());
						CellStyle copy_cellStyle = export_workbook.createCellStyle();
						copy_cellStyle.cloneStyleFrom(export_sheet.getRow(rowCopyIndex).getCell(cell.getCellIndex()).getCellStyle());
						copyCelll.setCellStyle(copy_cellStyle);
						copyCelll.setCellValue(cell.getCellValue());
					}
				}
			}
		}



		FileOutputStream out = new FileOutputStream("my_test_temp/sxssf.xlsx");
		export_workbook.write(out);
		out.close();
		// dispose of temporary files backing this workbook on disk
		export_workbook.dispose();
		export_workbook.close();
	}

	private List<CellReadCallBack> findCopyRow(List<CellReadCallBack> tempContentCollection, ListAndTableEntity listAndTableEntity,int cur_rowIndex) {
		List<?> list = listAndTableEntity.getDataList();
		for(Object list_obj : list) {
			if(list_obj != null) {
				List<?> list_obj_list = (List<?>)list_obj;
				for(Object obj : list_obj_list) {
					Set<String> keySet = new HashSet<>();
					if(obj instanceof Map) {
						Map<?,?> map = (Map<?,?>)obj;
						for(Entry<?, ?> entry : map.entrySet()) {
							String key_word = "list."+entry.getKey().toString();
							keySet.add(key_word);
						}
					}else {
						Field[] v_obj_field_arr = obj.getClass().getDeclaredFields();
						for (Field field : v_obj_field_arr) {
							String keyName = field.getName();
							String keyName_word = "list."+keyName;
							keySet.add(keyName_word);
						}
					}
					Iterator<String> keyIt = keySet.iterator();
					while(keyIt.hasNext()) {
						String pl_key = keyIt.next();
						for(CellReadCallBack cell : tempContentCollection) {
							if(cell.getCellValue()!= null && cell.getCellValue().startsWith(this.defaultPlaceholderPrefix) && cell.getCellValue().contains(pl_key)) {
								if(cur_rowIndex >= cell.getRowIndex() && cur_rowIndex - cell.getRowIndex()+1 < list_obj_list.size()) {
									List<CellReadCallBack> rtCellList = new ArrayList<>();
									for(CellReadCallBack cellrt : tempContentCollection) {
										if(cellrt.getRowIndex() == cell.getRowIndex()) {
											rtCellList.add(cellrt);
										}
									}
									return rtCellList;
								}
							}
						}
					}
				}
			}
		}
		return null;
	}


	private void copyTempWorkbookToSxssfWorkbook(XSSFWorkbook tempWorkbook, SXSSFWorkbook export_workbook) {
		int temp_sheetEnd = tempWorkbook.getNumberOfSheets();
		for (int temp_sheetIndex = 0; temp_sheetIndex < temp_sheetEnd; temp_sheetIndex++) {

			Sheet temp_sheet = tempWorkbook.getSheetAt(temp_sheetIndex);
			SXSSFSheet export_sheet = export_workbook.createSheet(temp_sheet.getSheetName());
			List<CellRangeAddress> temp_cellRangeAddressList = temp_sheet.getMergedRegions();
			for(CellRangeAddress temp_address_merge : temp_cellRangeAddressList) {
				export_sheet.addMergedRegion(temp_address_merge);
			}

			SXSSFDrawing patriarch = export_sheet.createDrawingPatriarch();

			Map<String, List<Object>> imgMap = new HashMap<String, List<Object>>();
			List<POIXMLDocumentPart> list = ((XSSFSheet)temp_sheet).getRelations();
			for (POIXMLDocumentPart part : list) {
				if (part instanceof XSSFDrawing) {
					XSSFDrawing drawing = (XSSFDrawing) part;
					List<XSSFShape> shapes = drawing.getShapes();
					for (XSSFShape shape : shapes) {
						XSSFPicture picture = (XSSFPicture) shape;
						XSSFClientAnchor anchor = picture.getPreferredSize();
						CTMarker marker = anchor.getFrom();
						String key = temp_sheetIndex + "-" + marker.getRow() + "-" + marker.getCol();
						imgMap.put(key, Arrays.asList(picture.getPictureData().getData(),anchor,picture.getPictureData().getPictureType()));
					}
				}
			}

			int temp_rowEnd = temp_sheet.getPhysicalNumberOfRows();
			for (int temp_rowIndex = 0; temp_rowIndex < temp_rowEnd; temp_rowIndex++) {
				Row temp_row = temp_sheet.getRow(temp_rowIndex);
				SXSSFRow exprot_row = export_sheet.createRow(temp_rowIndex);
				exprot_row.setHeight(temp_row.getHeight());

				int temp_cellEnd = temp_row.getPhysicalNumberOfCells();
				int temp_cell_last_num = temp_row.getLastCellNum();
				temp_cellEnd = temp_cellEnd>temp_cell_last_num?temp_cellEnd:temp_cell_last_num;
				for (int temp_cellIndex = 0; temp_cellIndex < temp_cellEnd; temp_cellIndex++) {
					export_sheet.setColumnWidth(temp_cellIndex, temp_sheet.getColumnWidth(temp_cellIndex));

					String img_key = temp_sheetIndex+"-"+temp_rowIndex+"-"+temp_cellIndex;
					if(imgMap.get(img_key) != null) {
						//XSSFClientAnchor anchor_export = new XSSFClientAnchor(0, 0, 0, 0, Integer.parseInt(img_key.split("-")[2]), Integer.parseInt(img_key.split("-")[1]), Integer.parseInt(img_key.split("-")[2]) + 1, Integer.parseInt(img_key.split("-")[1]) + 1);

						int picIndex = export_workbook.addPicture((byte[])imgMap.get(img_key).get(0), (int)imgMap.get(img_key).get(2));
						patriarch.createPicture((XSSFClientAnchor) imgMap.get(img_key).get(1), picIndex);
					}

					Cell temp_cell = temp_row.getCell(temp_cellIndex);
					if(temp_cell != null) {
						String temp_cell_value = PoiUtils.getCellValue(temp_cell);
						SXSSFCell export_cell = exprot_row.createCell(temp_cellIndex,temp_cell.getCellType());
						CellStyle export_cellStyle = export_workbook.createCellStyle();
						export_cellStyle.cloneStyleFrom(temp_cell.getCellStyle());
						export_cell.setCellValue(temp_cell_value);
						export_cell.setCellStyle(export_cellStyle);
					}
				}
			}
		}

	}

	/**
	 * 解析${占位符表达式
	 * @param cell_value
	 * @param dataList
	 * @return
	 * @throws Throwable 
	 */
	private String resolveCellValue(String cell_value, ListAndTableEntity listAndTableEntity,List<CellReadCallBack> tempContentCollection,int curRowIndex) throws Throwable {
		if(cell_value != null) {
			int start = cell_value.indexOf(this.defaultPlaceholderPrefix);
			if (start == -1) {
				return cell_value;
			}
			int listIndex = ExportUtils.findListIndexByKey(tempContentCollection, cell_value, curRowIndex);
			int end = cell_value.indexOf(this.defaultPlaceholderSuffix);
			String cell_value_word = cell_value.substring(cell_value.indexOf(this.defaultPlaceholderPrefix)+this.defaultPlaceholderPrefix.length(),cell_value.lastIndexOf(this.defaultPlaceholderSuffix));
			List<String> keyList = ExportUtils.getExpAllKeys(cell_value_word);

			Map<String, Object> expMap = new HashMap<>();

			for(String key : keyList) {
				List<?> v_list = listAndTableEntity.getDataList();
				if(v_list != null && v_list.size() > 0) {
					for(Object v_obj : v_list) {
						if(v_obj != null) {
							List<?> v_obj_list = (List<?>)v_obj;
							if(v_obj_list.size() > 0 && v_obj_list.size() > listIndex) {
								System.err.println("listIndex="+listIndex);
								Object list_obj = v_obj_list.get(listIndex);
								if(list_obj instanceof Map) {
									Map<?,?> list_obj_map = (Map<?, ?>) list_obj;
									for(Entry<?, ?> entry : list_obj_map.entrySet()) {
										String keyName_word = "list."+entry.getKey();
										if(keyName_word.equals(key)) {
											expMap.put(keyName_word, entry.getValue());
										}
									}
								}else {
									Field[] v_obj_field_arr = list_obj.getClass().getDeclaredFields();
									for (Field field : v_obj_field_arr) {
										field.setAccessible(true);
										String keyName = field.getName();
										String keyName_word = "list."+keyName;
										if(keyName_word.equals(key)) {
											expMap.put(keyName_word, field.get(v_obj_list.get(listIndex)));
										}
									}
								}
							}
						}
					}
				}

				List<?> tabList = listAndTableEntity.getTableList();
				if(tabList != null && tabList.size() > 0) {
					for(Object v_obj : tabList) {
						if(v_obj != null) {
							if(v_obj instanceof Map) {
								Map<?,?> list_obj_map = (Map<?, ?>) v_obj;
								for(Entry<?, ?> entry : list_obj_map.entrySet()) {
									String keyName_word = "list."+entry.getKey();
									if(keyName_word.equals(key)) {
										expMap.put(keyName_word, entry.getValue());
									}
								}
							}else {
								Field[] v_obj_field_arr = v_obj.getClass().getDeclaredFields();
								for (Field field : v_obj_field_arr) {
									field.setAccessible(true);
									String keyName = field.getName();
									String keyName_word = "tab."+keyName;
									if(keyName_word.equals(key)) {
										expMap.put(keyName_word, field.get(v_obj));
									}
								}
							}
						}
					}
				}

			}
			Object newCellValue = null;
			if(expMap.size() == keyList.size()) {
				newCellValue = AviatorEvaluator.compile(cell_value_word).execute(expMap);
			}
			if(newCellValue != null) {
				return new StringBuilder(cell_value).replace(start, end, cell_value_word).replace(start, end, newCellValue.toString()).toString();
			}
		}
		return cell_value;
	}

	public int findEndSheetIndex(Workbook workbook) {
		return workbook.getNumberOfSheets();
	}
	public Row findRow(Workbook workbook,int sheetIndex,int rowIndex) {
		return workbook.getSheetAt(sheetIndex).getRow(rowIndex);
	}

}
