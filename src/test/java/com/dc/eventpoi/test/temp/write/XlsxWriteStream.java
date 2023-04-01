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

/**
 * https://poi.apache.org/components/spreadsheet/how-to.html#sxssf
 * @author beijing-penguin
 *
 */
public class XlsxWriteStream {


	public void exportExcel(byte[] tempExcelBtye,List<Object> dataList) throws Throwable {
		
		SXSSFWorkbook export_workbook = new SXSSFWorkbook(-1);//关闭自动刷到磁盘，按行数计算
		export_workbook.setCompressTempFiles(true);//临时文件将被压缩

		XSSFWorkbook tempWorkbook = new XSSFWorkbook(new ByteArrayInputStream(tempExcelBtye));

		this.copyTempWorkbookToSxssfWorkbook(tempWorkbook,export_workbook);

		//开始写入数据
		int export_sheetEnd = tempWorkbook.getNumberOfSheets();
		for (int export_sheetIndex = 0; export_sheetIndex < export_sheetEnd; export_sheetIndex++) {
			Sheet export_sheet = tempWorkbook.getSheetAt(export_sheetIndex);

			List<Object> placeholderList = new ArrayList<>();
			int tempValueListIndex = 0;
			
			
			int export_rowEnd = export_sheet.getPhysicalNumberOfRows();
			for (int export_rowIndex = 0; export_rowIndex < export_rowEnd; export_rowIndex++) {
				Row export_row = export_sheet.getRow(export_rowIndex);

				int export_cellEnd = export_row.getPhysicalNumberOfCells();
				for (int export_cellIndex = 0; export_cellIndex < export_cellEnd; export_cellIndex++) {

					Cell export_cell = export_row.getCell(export_cellIndex);
					if(export_cell != null) {
						String export_cell_value = PoiUtils.getCellValue(export_cell);
						replaceCellValue(export_cell_value, dataList,tempValueListIndex);
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
						//String new_temp_cell_value = replaceCellValue(temp_cell_value,listAndTableSet);
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
     * 默认前缀占位符
     */
    public String defaultPlaceholderPrefix = "${";

    /**
     * 默认后缀占位符
     */
    public String defaultPlaceholderSuffix = "}";
	
    /**
     * 替换值中的占位符
     * @param cell_value
     * @param dataList
     * @return
     * @throws Throwable 
     */
	private String replaceCellValue(String cell_value, List<Object> dataList,int tempValueListIndex) throws Throwable {
		if(dataList.size() > 0 && cell_value != null) {
				
			int start = cell_value.indexOf(this.defaultPlaceholderPrefix);
	        if (start == -1) {
	            return cell_value;
	        }
	        int valueIndex = 0;
	        StringBuilder result = new StringBuilder(cell_value);
	        while (start != -1) {
	            int end = result.indexOf(this.defaultPlaceholderSuffix);
	            String cell_key = result.substring(start, end+this.defaultPlaceholderSuffix.length());
	            
//	            for(Object dataObj : dataList) {
//	            	if(dataObj instanceof List) {
//	            		List<?> objList = (List<?>)dataObj;
//	            		if(objList!=null && objList.size() > 0) {
//	            			for(Object obj : objList) {
//	            				Field[] fields = obj.getClass().getDeclaredFields();
//	            				for(Field field : fields) {
//	            					field.getName()
//	            				}
//	            			}
//	            		}
//	            	}
//	            }
	            
	            String newData = "";
	            Map<String, Object> expMap = new HashMap<>();
	            
	            if(cell_key.startsWith("list.")) {
	            	String cell_key_word = cell_key.substring(cell_key.indexOf("list."),cell_key.length());
	            	for(Object data : dataList) {
	            		if(data instanceof List) {
	            			List<?> v_list = (List<?>)data;
	            			if(v_list.size() > 0) {
	            				Object dataListObj = v_list.get(tempValueListIndex);
	            				Field[] fields = dataListObj.getClass().getDeclaredFields();
	            			    for (Field field : fields) {
	            			        field.setAccessible(true);
	            			        String keyName = field.getName();
	            			        if(cell_key.contains(keyName) ) {
	            			        	Object value = field.get(dataListObj);
	                			        if (value != null) {
	                			        	expMap.put(cell_key, value);
	                			        }
	            			        }
	            			    }
	            			}
	            		}
	            	}
	            } else if(cell_key.startsWith("tab.")) {
	            	//String cell_key_word = cell_key.substring(cell_key.indexOf("T:"),cell_key.length());
	            	for(Object data : dataList) {
	            		if(data instanceof List) {
	            		}else if (data instanceof Map) {
	            			Map<?,?> map = (Map<?,?>)data;
	            			for(Entry<?,?> mapData : map.entrySet()) {
	            				if(cell_key.con)
	            			}
	            		}else {
	        				Field[] fields = data.getClass().getDeclaredFields();
	        			    for (Field field : fields) {
	        			        field.setAccessible(true);
	        			        String keyName = field.getName();
	        			        if(cell_key.contains(keyName) ) {
	        			        	Object value = field.get(data);
	            			        if (value != null) {
	            			        	expMap.put(keyName, value);
	            			        }
	        			        }
	        			    }
	            		}
	            	}
	            }
	            newData = AviatorEvaluator.compile(cell_key).execute(expMap).toString();
	            result.replace(start, end + this.defaultPlaceholderSuffix.length(), newData);
	            start = result.indexOf(this.defaultPlaceholderPrefix, start + newData.length());
	        }
			return null;
		}
		return null;
	}

	public int findEndSheetIndex(Workbook workbook) {
		return workbook.getNumberOfSheets();
	}
	public Row findRow(Workbook workbook,int sheetIndex,int rowIndex) {
		return workbook.getSheetAt(sheetIndex).getRow(rowIndex);
	}

}
