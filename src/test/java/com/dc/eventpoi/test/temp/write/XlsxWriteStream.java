package com.dc.eventpoi.test.temp.write;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

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
import com.dc.eventpoi.test.temp.read.XlsxReadStream;
import com.googlecode.aviator.AviatorEvaluator;

/**
 * https://poi.apache.org/components/spreadsheet/how-to.html#sxssf
 * @author beijing-penguin
 *
 */
public class XlsxWriteStream {


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
			Collection<List<CellReadCallBack>> tempContentCollection = readTemp.doRead().values();
			
			
			SXSSFSheet export_sheet = export_workbook.getSheetAt(export_sheetIndex);
			
			int export_rowEnd = export_sheet.getPhysicalNumberOfRows();
			for (int export_rowIndex = 0; export_rowIndex < export_rowEnd; export_rowIndex++) {
				Row export_row = export_sheet.getRow(export_rowIndex);

				int export_cellEnd = export_row.getPhysicalNumberOfCells();
				for (int export_cellIndex = 0; export_cellIndex < export_cellEnd; export_cellIndex++) {

					Cell export_cell = export_row.getCell(export_cellIndex);
					if(export_cell != null) {
						String export_cell_value = PoiUtils.getCellValue(export_cell);
						String newCellValue = resolveCellValue(export_cell_value, listAndTableEntity,tempContentCollection,export_rowIndex);
						System.err.println(newCellValue);
						export_cell.setCellValue(newCellValue);
					}
				}
				
				//是否需要复制前一行
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
     * 解析${占位符表达式
     * @param cell_value
     * @param dataList
     * @return
     * @throws Throwable 
     */
	private String resolveCellValue(String cell_value, ListAndTableEntity listAndTableEntity,Collection<List<CellReadCallBack>> tempContentCollection,int curRowIndex) throws Throwable {
		if(cell_value != null) {
				
			int start = cell_value.indexOf(this.defaultPlaceholderPrefix);
	        if (start == -1) {
	            return cell_value;
	        }
	        String cell_value_word = cell_value.substring(cell_value.indexOf(this.defaultPlaceholderPrefix)+this.defaultPlaceholderPrefix.length(),cell_value.lastIndexOf(this.defaultPlaceholderSuffix));
	        Map<String, Object> expMap = new HashMap<>();
	        int tempIndex = 0;
	        while(cell_value_word.substring(tempIndex).contains("list.") || cell_value_word.substring(tempIndex).contains("tab.")) {
	            List<?> v_list = listAndTableEntity.getDataList();
	            if(v_list != null && v_list.size() > 0) {
	            	//找到这一列的对应的集合
	            	for(Object v_obj : v_list) {
	            		List<?> v_obj_list = (List<?>)v_obj;
	            		if(v_obj_list.size() > 0) {
	            			int listIndex = ExportUtils.findListIndexByKey(tempContentCollection, cell_value, curRowIndex);
	            			Field[] v_obj_field_arr = v_obj_list.get(listIndex).getClass().getDeclaredFields();
	            			for (Field field : v_obj_field_arr) {
	            		        field.setAccessible(true);
	            		        String keyName = field.getName();
	            		        String keyName_word = "list."+keyName;
	            		        int wordIndex = cell_value_word.indexOf(keyName_word);
	            		        if(wordIndex != -1 && (keyName_word.length() == cell_value_word.length() || ExportUtils.isWordKeyPart(cell_value_word.substring(wordIndex, wordIndex+1)))) {
	            		        	expMap.put(keyName_word, field.get(v_obj_list.get(listIndex)));
	            		        	tempIndex = keyName_word.length()-1;
	            		        	break;
	            		        }
	            		    }
	            		}
	            	}
	            }
	            tempIndex++;
	        }
	        if(expMap.size() > 0) {
	        	Object newCellValue = AviatorEvaluator.compile(cell_value_word).execute(expMap);
	        	return newCellValue==null?"":newCellValue.toString();
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
