package com.dc.eventpoi.test.temp.write;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
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

		//this.copyTempWorkbookToSxssfWorkbook(tempWorkbook,export_workbook);

		//开始写入数据
		int export_sheetEnd = tempWorkbook.getNumberOfSheets();
		for (int export_sheetIndex = 0; export_sheetIndex < export_sheetEnd; export_sheetIndex++) {

			Sheet temp_sheet = tempWorkbook.getSheetAt(export_sheetIndex);
			SXSSFSheet export_sheet = export_workbook.createSheet(temp_sheet.getSheetName());

			int temp_row_End = temp_sheet.getPhysicalNumberOfRows();
			int list_row_index = 0;
			for (int temp_row_index = 0; temp_row_index < temp_row_End; temp_row_index++) {
				Row temp_row = temp_sheet.getRow(temp_row_index);

				int temp_cellEnd = temp_sheet.getRow(0).getPhysicalNumberOfCells();
				int temp_cell_last_num = temp_sheet.getRow(0).getLastCellNum();
				temp_cellEnd = temp_cellEnd>temp_cell_last_num?temp_cellEnd:temp_cell_last_num;

				List<Cell> allTempCell = new ArrayList<Cell>();
				for (int temp_cell_index = 0; temp_cell_index < temp_cellEnd; temp_cell_index++) {
					Cell temp_cell = temp_row.getCell(temp_cell_index);
					allTempCell.add(temp_cell);
				}

				Map<Integer,CellStyle> cacheCellStyleMap = new HashMap<>();
				while(true) {
					SXSSFRow export_row = export_sheet.createRow(list_row_index+temp_row_index);
					export_row.setHeight(temp_row.getHeight());

					if(export_row.getRowNum() % 100 == 0) {
						export_sheet.flushRows(100);
					}

					//遍历当前总列数
					for (int export_cell_index = 0; export_cell_index < temp_cellEnd; export_cell_index++) {
						export_sheet.setColumnWidth(export_cell_index, temp_sheet.getColumnWidth(export_cell_index));
						Cell tempCell = allTempCell.get(export_cell_index);
						if(tempCell != null) {
							SXSSFCell export_cell = export_row.createCell(export_cell_index, tempCell.getCellType());
							CellStyle cellStyle = cacheCellStyleMap.get(export_cell_index);
							if(cellStyle == null) {
								cellStyle = export_workbook.createCellStyle();
								cellStyle.cloneStyleFrom(tempCell.getCellStyle());
								cacheCellStyleMap.put(export_cell_index, cellStyle);
							}
							export_cell.setCellStyle(cellStyle);
							String temp_cell_value = PoiUtils.getCellValue(tempCell);
							//解析temp_cell_value	
							String new_temp_cell_value = temp_cell_value;
							if(temp_cell_value.contains(this.defaultPlaceholderPrefix)) {
								String old_temp_cell_value = temp_cell_value;
								while(old_temp_cell_value.indexOf(defaultPlaceholderSuffix)!= -1) {
									//取出第一个${xxxx}表达式，xxxx
									String keyStr = old_temp_cell_value.substring(old_temp_cell_value.indexOf(defaultPlaceholderPrefix)+defaultPlaceholderPrefix.length(), old_temp_cell_value.indexOf(defaultPlaceholderSuffix));
									List<String> keyList = ExportUtils.getExpAllKeys(keyStr);
									Map<String, Object> expMap = new HashMap<>();

									for(String key : keyList) {
										List<?> v_list = listAndTableEntity.getDataList();
										if(v_list != null && v_list.size() > 0) {
											for(Object v_obj : v_list) {
												if(v_obj != null) {
													List<?> v_obj_list = (List<?>)v_obj;
													if(v_obj_list.size() > 0 && v_obj_list.size() > list_row_index) {
														Object list_obj = v_obj_list.get(list_row_index);
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
																	expMap.put(keyName_word, field.get(v_obj_list.get(list_row_index)));
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
										newCellValue = AviatorEvaluator.compile(keyStr).execute(expMap);
									}
									if(newCellValue != null) {
										new_temp_cell_value = new_temp_cell_value.replace(this.defaultPlaceholderPrefix+keyStr+this.defaultPlaceholderSuffix, newCellValue.toString());
									}

									old_temp_cell_value = old_temp_cell_value.substring(old_temp_cell_value.indexOf(defaultPlaceholderSuffix)+defaultPlaceholderSuffix.length());
								}
							}
							//解析完成
							//设置new_temp_cell_value
							export_cell.setCellValue(new_temp_cell_value);
						}
					}

					//判断是否还能继续创建行
					int l_r_index = list_row_index+1;
					boolean continueCreateRowFlag = false;
					List<?> list_list = listAndTableEntity.getDataList();
					for(Cell temp_cell : allTempCell) {
						if(temp_cell != null) {
							String temp_cell_value = PoiUtils.getCellValue(temp_cell);
							if(temp_cell_value != null && temp_cell_value.contains(this.defaultPlaceholderPrefix)) {
								for(Object list_obj : list_list) {
									if(list_obj != null) {
										List<?> dataList = (List<?>) list_obj;
										if(l_r_index < dataList.size()) {
											Object dataObj = dataList.get(l_r_index);
											if(dataObj instanceof Map) {
												Map<?, ?> dataMap = (Map<?, ?>) dataObj;
												for(Object keyObj : dataMap.keySet()) {
													if(temp_cell_value.contains("list."+keyObj.toString())) {
														continueCreateRowFlag = true;
													}
												}
											}else {
												Field[] v_obj_field_arr = dataObj.getClass().getDeclaredFields();
												for (Field field : v_obj_field_arr) {
													String keyName = field.getName();
													String keyName_word = "list."+keyName;
													if(temp_cell_value.contains(keyName_word)) {
														continueCreateRowFlag = true;
													}
												}
											}
										}
									}
								}
							}
						}
					}
					if(continueCreateRowFlag == false) {
						break;
					}else {
						list_row_index++;
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
}
