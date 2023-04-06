package com.dc.eventpoi.core;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFDrawing;
import org.apache.poi.xssf.streaming.SXSSFPicture;
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
	private String defaultPlaceholderPrefix = "${";

	/**
	 * 默认后缀占位符
	 */
	private String defaultPlaceholderSuffix = "}";

	/**
	 * 自动清除未匹配到的占位符
	 */
	private boolean autoClearPlaceholder = true;

	private Integer sheetIndex;

	public byte[] exportExcel(byte[] tempExcelBtye,ListAndTableEntity listAndTableEntity) throws Throwable {

		SXSSFWorkbook export_workbook = new SXSSFWorkbook(-1);//关闭自动刷到磁盘，按行数计算
		export_workbook.setCompressTempFiles(true);//临时文件将被压缩

		XSSFWorkbook tempWorkbook = new XSSFWorkbook(new ByteArrayInputStream(tempExcelBtye));

		Map<Object,Field[]> cacheObject = new HashMap<>();
		Map<String,Object> cacheExpMap = new HashMap<>();
		
		int sheetStart = 0;
        int sheetEnd = tempWorkbook.getNumberOfSheets();
        if (sheetIndex != null) {
            sheetStart = sheetIndex;
            sheetEnd = sheetIndex + 1;
        }
        
		//开始写入数据
		for (int export_sheetIndex = sheetStart; export_sheetIndex < sheetEnd; export_sheetIndex++) {
			
			Sheet temp_sheet = tempWorkbook.getSheetAt(export_sheetIndex);
			SXSSFSheet export_sheet = export_workbook.createSheet(temp_sheet.getSheetName());

			//添加合并行
			List<CellRangeAddress> temp_cellRangeAddressList = temp_sheet.getMergedRegions();
			for(CellRangeAddress temp_address_merge : temp_cellRangeAddressList) {
				export_sheet.addMergedRegion(temp_address_merge);
			}
			
			
			//提取出模板中的所有图片
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
						String key = export_sheetIndex + "-" + marker.getRow() + "-" + marker.getCol();
						imgMap.put(key, Arrays.asList(picture.getPictureData().getData(),anchor,picture.getPictureData().getPictureType()));
					}
				}
			}
			
			
			int temp_row_End = temp_sheet.getPhysicalNumberOfRows();
			int list_row_index = 0;
			for (int temp_row_index = 0; temp_row_index < temp_row_End; temp_row_index++) {
				
//				List<?> vvv_list = listAndTableEntity.getDataList();
//				if(vvv_list != null && vvv_list.size() > 0) {
//					for(Object vvv_obj : vvv_list) {
//						if(((List<?>)vvv_obj).size() <= list_row_index ) {
//							continueCreateRowFlag = false;
//							break;
//						}
//					}
//				}
				
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

					if(export_row.getRowNum() % 1000 == 0) {
						export_sheet.flushRows(1000);
					}

					//遍历当前总列数
					//boolean continueCreateRowFlag = false;
					for (int export_cell_index = 0; export_cell_index < temp_cellEnd; export_cell_index++) {
						
						//添加图片
						String img_key = export_sheetIndex+"-"+export_row.getRowNum()+"-"+export_cell_index;
						if(imgMap.get(img_key) != null) {
							//XSSFClientAnchor anchor_export = new XSSFClientAnchor(0, 0, 0, 0, Integer.parseInt(img_key.split("-")[2]), Integer.parseInt(img_key.split("-")[1]), Integer.parseInt(img_key.split("-")[2]) + 1, Integer.parseInt(img_key.split("-")[1]) + 1);

							int picIndex = export_workbook.addPicture((byte[])imgMap.get(img_key).get(0), (int)imgMap.get(img_key).get(2));
							patriarch.createPicture((XSSFClientAnchor) imgMap.get(img_key).get(1), picIndex);
						}
						
						export_sheet.setColumnWidth(export_cell_index, temp_sheet.getColumnWidth(export_cell_index));
						Cell tempCell = allTempCell.get(export_cell_index);
						if(tempCell != null) {
							SXSSFCell export_cell = export_row.createCell(export_cell_index, tempCell.getCellType());
							CellStyle cellStyle = cacheCellStyleMap.get(export_cell_index);
							if(cellStyle == null) {
								cellStyle = export_workbook.createCellStyle();
								cellStyle.cloneStyleFrom(tempCell.getCellStyle());
								cellStyle.setFillBackgroundColor(tempCell.getCellStyle().getFillBackgroundColorColor());
								cacheCellStyleMap.put(export_cell_index, cellStyle);
							}
							export_cell.setCellStyle(cellStyle);
							String temp_cell_value = PoiUtils.getCellValue(tempCell);
							//解析temp_cell_value	
							String new_temp_cell_value = temp_cell_value;
							byte[] image_bytes = null;
							if(temp_cell_value.contains(this.defaultPlaceholderPrefix)) {
								int tp_start_index = 0;
								while(new_temp_cell_value.indexOf(defaultPlaceholderPrefix,tp_start_index)!= -1) {
									//取出第一个${xxxx}表达式，xxxx
									String keyStr = new_temp_cell_value.substring(new_temp_cell_value.indexOf(defaultPlaceholderPrefix)+defaultPlaceholderPrefix.length(), new_temp_cell_value.indexOf(defaultPlaceholderSuffix));
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
																	//continueCreateRowFlag = true;
																	if(entry.getValue().getClass().getTypeName().equals("byte[]")) {
																		image_bytes = (byte[])entry.getValue();
																	}else {
																		expMap.put(keyName_word, entry.getValue());
																	}
																}
															}
														}else {
															Field[] v_obj_field_arr = list_obj.getClass().getDeclaredFields();
															for (Field field : v_obj_field_arr) {
																field.setAccessible(true);
																String keyName = field.getName();
																String keyName_word = "list."+keyName;
																if(keyName_word.equals(key)) {
																	//continueCreateRowFlag = true;
																	if(field.getType().getTypeName().equals("byte[]")) {
																		image_bytes = (byte[])field.get(v_obj_list.get(list_row_index));
																	}else {
																		expMap.put(keyName_word, field.get(v_obj_list.get(list_row_index)));
																	}
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
															String keyName_word = "tab."+entry.getKey();
															if(keyName_word.equals(key)) {
																if(entry.getValue().getClass().getTypeName().equals("byte[]")) {
																	image_bytes = (byte[])entry.getValue();
																}else {
																	expMap.put(keyName_word, entry.getValue());
																}
															}
														}
													}else {
														Field[] v_obj_field_arr = cacheObject.get(v_obj);
														if(v_obj_field_arr == null) {
															v_obj_field_arr = v_obj.getClass().getDeclaredFields();
															cacheObject.put(v_obj, v_obj_field_arr);
														}
														
														//Field[] v_obj_field_arr = v_obj.getClass().getDeclaredFields();
														for (Field field : v_obj_field_arr) {
															field.setAccessible(true);
															String keyName = field.getName();
															String keyName_word = "tab."+keyName;
															if(keyName_word.equals(key)) {
																if(field.getType().getTypeName().equals("byte[]")) {
																	image_bytes = (byte[])field.get(v_obj);
																}else {
																	expMap.put(keyName_word, field.get(v_obj));
																}
															}
														}
													}
												}
											}
										}
									}

									String pl_key_str = this.defaultPlaceholderPrefix+keyStr+this.defaultPlaceholderSuffix;
									Object newCellValue = null;
									if(expMap.size() == keyList.size()) {
										//newCellValue = AviatorEvaluator.compile(keyStr).execute(expMap);
										if(expMap.size() == 1 && temp_cell_value.contains(this.defaultPlaceholderPrefix+expMap.keySet().iterator().next()+this.defaultPlaceholderSuffix)) {
											newCellValue = expMap.values().iterator().next();
										}else {
											String key = keyStr + expMap;
											if(cacheExpMap.containsKey(key)) {
												newCellValue = cacheExpMap.get(key);
											}else {
												newCellValue = AviatorEvaluator.compile(keyStr).execute(expMap);
												cacheExpMap.put(key, newCellValue);
											}
											//newCellValue = AviatorEvaluator.compile(keyStr).execute(expMap);
										}
									}
									
									if(newCellValue != null) {
										tp_start_index = tp_start_index + new_temp_cell_value.indexOf(this.defaultPlaceholderPrefix)+newCellValue.toString().length();
										new_temp_cell_value = new_temp_cell_value.replace(pl_key_str, newCellValue.toString());
									}else {
										if(autoClearPlaceholder == true) {
											new_temp_cell_value = new_temp_cell_value.replace(pl_key_str, "");
										}else {
											tp_start_index = tp_start_index + pl_key_str.length();
										}
									}
								}
							}
							//解析完成
							//设置new_temp_cell_value
							if(image_bytes == null) {
								export_cell.setCellValue(new_temp_cell_value);
							}else {
			        	    	// 获取单元格宽度和高度，单位都是像素
			        	    	double cellWidth = temp_sheet.getColumnWidthInPixels(export_cell_index);
			        	    	//double cellHeight = temp_row.getHeightInPoints() / 72 * 96;// getHeightInPoints()方法获取的是点（磅），就是excel设置的行高，1英寸有72磅，一般显示屏一英寸是96个像素
				
								//设置图片
								XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, export_cell_index, temp_row_index+list_row_index, export_cell_index+1, temp_row_index+list_row_index+1);
								//anchor.setAnchorType(AnchorType.DONT_MOVE_AND_RESIZE); // 设置为不拉伸
								//anchor.setAnchorType(AnchorType.MOVE_AND_RESIZE);
                                int picIndex = export_workbook.addPicture(image_bytes, HSSFWorkbook.PICTURE_TYPE_JPEG);
                                SXSSFPicture pic = patriarch.createPicture(anchor, picIndex);
                                
                                int imageWidth = pic.getImageDimension().width;
                                if (imageWidth > cellWidth) {
                    	    		double scaleX = cellWidth / imageWidth;// 最终图片大小与单元格宽度的比例
                    	    		// 最终图片大小与单元格高度的比例
                    	    		// 说一下这个比例的计算方式吧：( imageHeight / imageWidth ) 是原图高于宽的比值，则 ( width * ( imageHeight / imageWidth ) ) 就是最终图片高的比值，
                    	    		// 那 ( width * ( imageHeight / imageWidth ) ) / cellHeight 就是所需比例了
                    		    	//double scaleY = ( cellWidth * ( imageHeight / imageWidth ) ) / cellHeight;
                    		    	pic.resize(scaleX, 1);
                    	    	}
							}
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


		tempWorkbook.close();
        ByteArrayOutputStream byteStream = new ByteArrayOutputStream();
        export_workbook.write(byteStream);
        byteStream.flush();
        byteStream.close();
        export_workbook.close();
        export_workbook.dispose();
        return byteStream.toByteArray();
	}


	public String getDefaultPlaceholderPrefix() {
		return defaultPlaceholderPrefix;
	}


	public void setDefaultPlaceholderPrefix(String defaultPlaceholderPrefix) {
		this.defaultPlaceholderPrefix = defaultPlaceholderPrefix;
	}


	public String getDefaultPlaceholderSuffix() {
		return defaultPlaceholderSuffix;
	}


	public void setDefaultPlaceholderSuffix(String defaultPlaceholderSuffix) {
		this.defaultPlaceholderSuffix = defaultPlaceholderSuffix;
	}


	public boolean isAutoClearPlaceholder() {
		return autoClearPlaceholder;
	}


	public void setAutoClearPlaceholder(boolean autoClearPlaceholder) {
		this.autoClearPlaceholder = autoClearPlaceholder;
	}


	public Integer getSheetIndex() {
		return sheetIndex;
	}


	public void setSheetIndex(Integer sheetIndex) {
		this.sheetIndex = sheetIndex;
	}
}
