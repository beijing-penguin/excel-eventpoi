///**
// * EventExcelHelper.java
// */
//package com.dc.eventpoi;
//
//import java.io.ByteArrayInputStream;
//import java.io.ByteArrayOutputStream;
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileNotFoundException;
//import java.io.InputStream;
//import java.lang.reflect.Field;
//import java.lang.reflect.Modifier;
//import java.nio.file.Files;
//import java.nio.file.Paths;
//import java.util.ArrayList;
//import java.util.Arrays;
//import java.util.Collection;
//import java.util.HashMap;
//import java.util.Iterator;
//import java.util.List;
//import java.util.Map;
//import java.util.Map.Entry;
//
//import org.apache.commons.lang3.reflect.FieldUtils;
//import org.apache.poi.hssf.usermodel.HSSFSheet;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.CellStyle;
//import org.apache.poi.ss.usermodel.CellType;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.ss.usermodel.WorkbookFactory;
//import org.apache.poi.ss.util.CellRangeAddress;
//import org.apache.poi.xssf.streaming.SXSSFCell;
//import org.apache.poi.xssf.streaming.SXSSFDrawing;
//import org.apache.poi.xssf.streaming.SXSSFRow;
//import org.apache.poi.xssf.streaming.SXSSFSheet;
//import org.apache.poi.xssf.streaming.SXSSFWorkbook;
//import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//import com.alibaba.fastjson.JSON;
//import com.dc.eventpoi.core.PoiUtils;
//import com.dc.eventpoi.core.entity.ExcelCell;
//import com.dc.eventpoi.core.entity.ExcelRow;
//import com.dc.eventpoi.core.entity.ExportExcelCell;
//import com.dc.eventpoi.core.enums.FileType;
//import com.dc.eventpoi.core.inter.CellStyleCallBack;
//import com.dc.eventpoi.core.inter.ExcelEventStream;
//import com.dc.eventpoi.core.inter.RowCallBack;
//import com.dc.eventpoi.core.inter.SheetCallBack;
//
///**
// * excel操作
// *
// * @author beijing-penguin
// */
//public class ExcelHelperBuild {
//
//	
//	private Integer importSheetIndex;
//	private String importSheetName;
//	private byte[] templateFileByte;
//	private Object[] tableList;
//	private Object[] dataList;
//	
//	public static ExcelHelperBuild build() {
//		return new ExcelHelperBuild();
//	}
//	
//	public ExcelHelperBuild setImportSheetIndex(Integer importSheetIndex) {
//		this.importSheetIndex = importSheetIndex;
//		return this;
//	}
//	
//	public ExcelHelperBuild setImportSheetName(String importSheetName) {
//		this.importSheetName = importSheetName;
//		return this;
//	}
//	
//	public ExcelHelperBuild setTemplateFileByte(File filePath) throws Exception {
//	    this.templateFileByte = Files.readAllBytes(Paths.get(filePath.toURI()));
//		return this;
//	}
//	public ExcelHelperBuild setTemplateFileByte(InputStream fileInputStream) throws Exception {
//	    this.templateFileByte = PoiUtils.inputStreamToByte(fileInputStream);
//		return this;
//	}
//	
//	public ExcelHelperBuild setTableList(Object...tabData) {
//	    this.tableList = tabData;
//        return this;
//    }
//	
//	public ExcelHelperBuild setDataList(Object[] dataList) {
//	    this.dataList = dataList;
//        return this;
//    }
//	public ExcelHelperBuild setDataList(List<?> dataList) {
//        this.dataList = dataList.toArray();
//        return this;
//    }
//	
//	public byte[] export() throws Exception {
//		Workbook tempWorkbook = null;
//        FileType fileType = PoiUtils.judgeFileType(new ByteArrayInputStream(templateFileByte));
//        if (fileType == FileType.XLSX) {
//        	tempWorkbook = new XSSFWorkbook(new ByteArrayInputStream(templateFileByte));
//        } else {
//        	tempWorkbook = (HSSFWorkbook) WorkbookFactory.create(new ByteArrayInputStream(templateFileByte));
//        }
//        
//        SXSSFWorkbook exportWorkbook = new SXSSFWorkbook();
//        int sheetEnd = tempWorkbook.getNumberOfSheets();
//        for (int sheetIndex = 0; sheetIndex < sheetEnd; sheetIndex++) {
//            
//        	SXSSFSheet exportSheet = exportWorkbook.createSheet(tempWorkbook.getSheetName(sheetIndex));
//        	SXSSFDrawing exportPatriarch = (SXSSFDrawing) exportSheet.createDrawingPatriarch();
//        	
//        	Sheet tempSheet = tempWorkbook.getSheetAt(sheetIndex);
//            int sheetMergerCount = tempSheet.getNumMergedRegions();
//            int rowNum = tempSheet.getPhysicalNumberOfRows();
//            Map<Integer,Row> newAddRow = new HashMap<Integer,Row>();
//            for (int rowIndex = 0; rowIndex < rowNum; rowIndex++) {
//                for (int ii = 0; ii < sheetMergerCount; ii++) {
//                    CellRangeAddress mergedRegionAt = tempSheet.getMergedRegion(ii);
//                    if (mergedRegionAt.getFirstRow() == rowIndex) {
//                        mergedRegionAt.setFirstRow(mergedRegionAt.getFirstRow() + newAddRow.size());
//                        mergedRegionAt.setLastRow(mergedRegionAt.getLastRow() + newAddRow.size());
//                        exportSheet.addMergedRegion(mergedRegionAt);
//                    }
//                }
//            	Row tempRow = tempSheet.getRow(rowIndex);
//            	Row exportRow = exportSheet.createRow(rowIndex);
//            	int tempCellNum = tempRow.getPhysicalNumberOfCells();
//            	for (int cellIndex = 0; cellIndex < tempCellNum; cellIndex++) {
//            		Cell tempCell = tempRow.getCell(cellIndex);
//            		CellStyle tempCellStyle = tempCell.getCellStyle();
//            		String tempCellValue = PoiUtils.getCellValue(tempCell);
//            		
//            		Cell exportCell = exportRow.createCell(cellIndex);
//            		exportCell.setCellStyle(tempCellStyle);
//            		
//            		if (tempCellValue != null && tempCellValue.contains("${")) {
//            		    //解析出字段明
//            			String keyName = tempCellValue.substring(tempCellValue.indexOf("${") + 2, tempCellValue.lastIndexOf("}"));
//            			//解析出包含{得整个占位符
//                        String excelFieldSrcKeyword = tempCellValue.substring(tempCellValue.indexOf("${"), tempCellValue.lastIndexOf("}") + 1);
//                        
//                        boolean isFind = false;
//                        //首先在table结构中寻找匹配
//                        if(tableList !=null && tableList.length > 0) {
//                        	for (Object tabObj : tableList) {
//                        	    Field field = FieldUtils.getField(tabObj.getClass(), keyName, true);
//                        	    Object fieldValue = null;
//                        	    if (field != null && (fieldValue=field.get(tabObj)) != null) {
//                        	        //判断是不是 图片数据
//                        	        if (fieldValue instanceof byte[]) {
//                                        if (PoiUtils.getImageType((byte[]) fieldValue) != null) {
//                                            XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, cellIndex, rowIndex, cellIndex + 1, rowIndex + 1);
//                                            int picIndex = exportWorkbook.addPicture((byte[]) fieldValue, HSSFWorkbook.PICTURE_TYPE_JPEG);
//                                            exportPatriarch.createPicture(anchor, picIndex);
//                                        } else {
//                                            String exportValue = tempCellValue.replace(excelFieldSrcKeyword, new String((byte[]) fieldValue));
//                                            exportCell.setCellValue(exportValue);
//                                        }
//                                    } else {
//                                        String exportValue = tempCellValue.replace(excelFieldSrcKeyword, String.valueOf(fieldValue));
//                                        exportCell.setCellValue(exportValue);
//                                    }
//                        	        isFind = true;
//                        	        break;
//                        	    }
//							}
//                        }
//                        
//                        if(isFind == false) {//表格没找到，继续在列表中查找
//                            Object dataObj = dataList[0];//只取第一行的值设置在当前cell
//                            Field field = FieldUtils.getField(dataObj.getClass(), keyName, true);
//                            Object fieldValue = null;
//                            if (field != null && (fieldValue = field.get(dataObj)) != null) {
//                                isFind = true;
//                                //设置当前整个列的数据
//                                for (int dataCellIndex = 0; dataCellIndex < dataList.length; dataCellIndex++) {
//                                    Object cellObj = dataList[dataCellIndex];
//                                    Field cellField = FieldUtils.getField(cellObj.getClass(), keyName, true);
//                                    Object fieldValue2 = cellField.get(dataObj);
//                                    
//                                    if(fieldValue2 == null) {
//                                        fieldValue2 = "";
//                                    }
//                                    int curRowIndex = rowIndex;
//                                    Cell curExportCell = exportCell;
//                                    if(dataCellIndex == 0) {
//                                        //String exportValue = tempCellValue.replace(excelFieldSrcKeyword, new String((byte[]) fieldValue2));
//                                        //curExportCell.setCellValue(exportValue);
//                                    }else {
//                                        curRowIndex = rowIndex+newAddRow.size()+1;
//                                        Row row = tempSheet.createRow(curRowIndex);
//                                        newAddRow.put(curRowIndex, row);
//                                        curExportCell = row.createCell(cellIndex,tempCell.getCellType());
//                                        curExportCell.setCellStyle(tempCellStyle);
//                                    }
//                                    
//                                    if (fieldValue2 instanceof byte[]) {
//                                        if (PoiUtils.getImageType((byte[]) fieldValue) != null) {
//                                            XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, cellIndex, curRowIndex, cellIndex + 1, curRowIndex + 1);
//                                            int picIndex = exportWorkbook.addPicture((byte[]) fieldValue2, HSSFWorkbook.PICTURE_TYPE_JPEG);
//                                            exportPatriarch.createPicture(anchor, picIndex);
//                                        } else {
//                                            String exportValue = tempCellValue.replace(excelFieldSrcKeyword, new String((byte[]) fieldValue2));
//                                            curExportCell.setCellValue(exportValue);
//                                        }
//                                    } else {
//                                        String exportValue = tempCellValue.replace(excelFieldSrcKeyword, String.valueOf(fieldValue2));
//                                        curExportCell.setCellValue(exportValue);
//                                    }
//                                }
//                            }
//                        }
//                        
//                        if(isFind == false) {//都没找到，设置为原Cell的值
//                            exportCell.setCellValue(String.valueOf(tempCellValue));
//                        }
//            		}else {//都没找到，设置为原Cell的值
//            		    exportCell.setCellValue(tempCellValue);
//            		    //PoiUtils.copyTempCell(tempCellStyle,tempCellValue,exportCell);
//            		}
//				}
//            }
//        }
//        
//        
//        tempWorkbook.close();
//        ByteArrayOutputStream byteStream = new ByteArrayOutputStream();
//        exportWorkbook.write(byteStream);
//        byteStream.flush();
//        byteStream.close();
//        exportWorkbook.close();
//        exportWorkbook.dispose();
//        return byteStream.toByteArray();
//	}
//    /**
//     * 导出表格 以及 列表数据
//     * 
//     * @param tempExcelBtye        模板文件流
//     * @param listAndTableDataList 包含列表数据集合 和 表格数据对象
//     * @param sheetIndex           sheetIndex
//     * @param sheetCallBack        sheetCallBack
//     * @param callBackCellStyle    callBackCellStyle
//     * @return byte[]
//     * @throws Exception Exception
//     */
//    public static byte[] exportExcel(byte[] tempExcelBtye, List<?> listAndTableDataList, Integer sheetIndex, SheetCallBack sheetCallBack, CellStyleCallBack callBackCellStyle) throws Exception {
//        boolean is_data_list = true;
//        for (int i = 0,len = listAndTableDataList.size(); i < len; i++) {
//			if(i == 5000) {
//				break;
//			}
//			Object dataObj = listAndTableDataList.get(i);
//			if(dataObj instanceof Collection) {
//				is_data_list = false;
//				break;
//			}
//		}
//    	
//    	Workbook workbook = null;
//        FileType fileType = PoiUtils.judgeFileType(new ByteArrayInputStream(tempExcelBtye));
//        if (fileType == FileType.XLSX) {
//            workbook = new XSSFWorkbook(new ByteArrayInputStream(tempExcelBtye));
//        } else {
//            workbook = (HSSFWorkbook) WorkbookFactory.create(new ByteArrayInputStream(tempExcelBtye));
//        }
//
//        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook();
//        
//        int sheetStart = 0;
//        int sheetEnd = workbook.getNumberOfSheets();
//        if (sheetIndex != null) {
//            sheetStart = sheetIndex;
//            sheetEnd = sheetIndex + 1;
//        }
//        for (int i = sheetStart; i < sheetEnd; i++) {
//            SXSSFSheet sxssSheet = sxssfWorkbook.createSheet(workbook.getSheetName(i));
//            if (sheetCallBack != null) {
//                sheetCallBack.callBack(sxssSheet);
//            }
//
//            SXSSFDrawing patriarch = (SXSSFDrawing) sxssSheet.createDrawingPatriarch();
//            Sheet xsssheet = workbook.getSheetAt(i);
//            int sheetMergerCount = xsssheet.getNumMergedRegions();
//
//            int rowNum = xsssheet.getPhysicalNumberOfRows();
//            int offset = 0;
//            int listCount = 0;
//            for (int j = 0; j < rowNum; j++) {
//                for (int ii = 0; ii < sheetMergerCount; ii++) {
//                    CellRangeAddress mergedRegionAt = xsssheet.getMergedRegion(ii);
//                    if (mergedRegionAt.getFirstRow() == j) {
//                        mergedRegionAt.setFirstRow(mergedRegionAt.getFirstRow() + offset - listCount);
//                        mergedRegionAt.setLastRow(mergedRegionAt.getLastRow() + offset - listCount);
//                        sxssSheet.addMergedRegion(mergedRegionAt);
//                    }
//                }
//
//                Row xssrow = xsssheet.getRow(j);
//                int xssCellNum = xssrow.getPhysicalNumberOfCells();
//                boolean breakFlag = false;
//
//                SXSSFRow sxssrow = sxssSheet.createRow(j + offset - listCount);
//                sxssrow.setHeight(xssrow.getHeight());
//
//                for (int k = 0; k < xssCellNum; k++) {
//                    final int temp_k = k;
//                    if (breakFlag) {
//                        break;
//                    }
//                    Cell xssCell = xssrow.getCell(k);
//                    sxssSheet.setColumnWidth(k, xsssheet.getColumnWidth(k));
//                    if (xssCell == null) {
//                    } else {
//                        boolean matchFlag = false;
//                        String xssCellValue = null;
//                        if(xssCell.getCellType() == CellType.NUMERIC) {
//                        	xssCellValue = String.valueOf(xssCell.getNumericCellValue());
//                        }else {
//                        	xssCellValue = xssCell.getStringCellValue();
//                        }
//                        if (xssCellValue != null && xssCellValue.contains("${")) {
//                            String keyName = xssCellValue.substring(xssCellValue.indexOf("${") + 2, xssCellValue.lastIndexOf("}"));
//                            String excelFieldSrcKeyword = xssCellValue.substring(xssCellValue.indexOf("${"), xssCellValue.lastIndexOf("}") + 1);
//
//                            for (Object dataObj : listAndTableDataList) {
//                                if (matchFlag) {
//                                    break;
//                                }
//                                if ((dataObj instanceof Collection) || is_data_list == true) {
//                                    List<?> dataList = null;
//                                    if(is_data_list == true) {
//                                    	dataList = listAndTableDataList;
//                                    }else {
//                                    	dataList = (List<?>) dataObj;
//                                    }
//                                    if (dataList.size() > 0) {
//                                        Object tempData = dataList.get(0);
//                                        if (FieldUtils.getField(tempData.getClass(), keyName, true) == null) {
//                                            continue;
//                                        }
//
//                                        List<ExportExcelCell> keyCellList = new ArrayList<ExportExcelCell>();
//                                        for (int kk = k; kk < xssCellNum; kk++) {
//                                            Cell xssCell_kk = xssrow.getCell(kk);
//                                            CellType type = xssCell_kk.getCellType();
//                                            
//                                            /*
//                                             * Color color = xssCell_kk.getCellStyle().getFillBackgroundColorColor();
//                                             * if(color != null) { System.err.println(((XSSFColor)color).getARGB());
//                                             * System.err.println(((XSSFColor)color).getRGB());
//                                             * System.err.println(((XSSFColor)color).getCTColor().xmlText());
//                                             * System.err.println(((XSSFColor)color).getCTColor().getRgb());
//                                             * System.err.println(((XSSFColor)color).getARGBHex()); }
//                                             */
//                                            //System.err.println(color);	
//                                            
//                                            CellStyle _sxssStyle = sxssfWorkbook.createCellStyle();
//                                            _sxssStyle.cloneStyleFrom(xssCell_kk.getCellStyle());
//                                            
//                                            ExportExcelCell ee = new ExportExcelCell((short) xssCell_kk.getColumnIndex(), xssCell_kk.getStringCellValue(), _sxssStyle);
//                                            ee.setCellType(type);
//                                            keyCellList.add(ee);
//                                        }
//                                        breakFlag = true;
//                                        matchFlag = true;
//                                        listCount++;
//                                        for (int y = 0,len=dataList.size(); y < len; y++) {
//                                            final int create_row_num = j + offset;
//                                            offset++;
//
//                                            Object srcData = dataList.get(y);
//                                            SXSSFRow sxssrow_y = sxssSheet.createRow(create_row_num);
//
//                                            sxssrow_y.setHeight(xssrow.getHeight());
//                                            for (int x = temp_k; x < xssCellNum; x++) {
//
//                                                ExportExcelCell curCell = null;
//                                                String vv = null;
//                                                for (ExportExcelCell exportCell : keyCellList) {
//                                                    if (exportCell.getIndex() == x) {
//                                                        curCell = exportCell;
//                                                        vv = exportCell.getValue();
//                                                        break;
//                                                    }
//                                                }
//                                                // curCell.getCellStyle().setFillForegroundColor(IndexedColors.AQUA.getIndex());
//                                                // curCell.getCellStyle().setFillPattern(FillPatternType.SOLID_FOREGROUND);
//                                                String _keyName = vv.substring(vv.indexOf("${") + 2, vv.lastIndexOf("}"));
//                                                Field field = FieldUtils.getField(srcData.getClass(), _keyName, true);
//                                                if (field != null && field.get(srcData) != null) {
//                                                    SXSSFCell _sxssCell = sxssrow_y.createCell(x, curCell.getCellType());
//                                                    if (callBackCellStyle != null) {
//                                                        callBackCellStyle.callBack(sxssSheet, _sxssCell, curCell.getCellStyle());
//                                                        _sxssCell.setCellStyle(curCell.getCellStyle());
//                                                    } else {
//                                                        _sxssCell.setCellStyle(curCell.getCellStyle());
//                                                    }
//
//                                                    Object value = field.get(srcData);
//                                                    if (value instanceof byte[]) {
//                                                        if (PoiUtils.getImageType((byte[]) value) != null) {
//                                                            //XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, x, sxssrow_y.getRowNum(), x + 1, sxssrow_y.getRowNum() + 1);
////                                                            int picIndex = sxssSheet.getWorkbook().addPicture((byte[]) value, HSSFWorkbook.PICTURE_TYPE_PNG);
////                                                            Drawing drawing = sxssSheet.getDrawingPatriarch();
////                                                            if (drawing == null) {
////                                                                drawing = sxssSheet.createDrawingPatriarch();
////                                                            }
////                                                            
////                                                            CreationHelper helper = sxssSheet.getWorkbook().getCreationHelper();
////                                                            ClientAnchor anchor = helper.createClientAnchor();
////                                                            anchor.setDx1(0);
////                                                            anchor.setDx2(0);
////                                                            anchor.setDy1(0);
////                                                            anchor.setDy2(0);
////                                                            anchor.setCol1(_sxssCell.getColumnIndex());
////                                                            anchor.setCol2(_sxssCell.getColumnIndex() + 1);
////                                                            anchor.setRow1(_sxssCell.getRowIndex());
////                                                            anchor.setRow2(_sxssCell.getRowIndex() + 1);
////                                                            anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
////                                                            
////                                                            drawing.createPicture(anchor, picIndex);
//                                                        	XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, x, sxssrow_y.getRowNum(), x + 1, sxssrow_y.getRowNum() + 1);
//                                                            int picIndex = sxssfWorkbook.addPicture((byte[]) value, HSSFWorkbook.PICTURE_TYPE_JPEG);
//                                                            patriarch.createPicture(anchor, picIndex);
//                                                        } else {
//                                                            _sxssCell.setCellValue(new String((byte[]) value));
//                                                        }
//                                                    } else {
//                                                        _sxssCell.setCellValue(String.valueOf(value));
//                                                    }
//                                                } else {
//                                                    SXSSFCell _sxssCell = sxssrow_y.createCell(x, curCell.getCellType());
//                                                    if (callBackCellStyle != null) {
//                                                        callBackCellStyle.callBack(sxssSheet, _sxssCell, curCell.getCellStyle());
//                                                        _sxssCell.setCellStyle(curCell.getCellStyle());
//                                                    } else {
//                                                        _sxssCell.setCellStyle(curCell.getCellStyle());
//                                                    }
//                                                    _sxssCell.setCellValue("");
//                                                }
//                                            }
//                                        }
//                                    }
//                                } else {
//                                    Field field = FieldUtils.getField(dataObj.getClass(), keyName, true);
//                                    if (field != null) {
//                                        matchFlag = true;
//                                        SXSSFCell sxssCell = sxssrow.createCell(k, xssCell.getCellType());
//                                        CellStyle _sxssStyle = sxssfWorkbook.createCellStyle();
//                                        if (callBackCellStyle != null) {
//                                            _sxssStyle.cloneStyleFrom(xssCell.getCellStyle());
//                                            sxssCell.setCellStyle(_sxssStyle);
//                                            callBackCellStyle.callBack(sxssSheet, sxssCell, _sxssStyle);
//                                        } else {
//                                            _sxssStyle.cloneStyleFrom(xssCell.getCellStyle());
//                                            sxssCell.setCellStyle(_sxssStyle);
//                                        }
//
//                                        Object value = field.get(dataObj);
//                                        if (value instanceof byte[]) {
//                                            if (PoiUtils.getImageType((byte[]) value) != null) {
//                                                XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, k, sxssrow.getRowNum(), k + 1, sxssrow.getRowNum() + 1);
//                                                int picIndex = sxssfWorkbook.addPicture((byte[]) value, HSSFWorkbook.PICTURE_TYPE_JPEG);
//                                                patriarch.createPicture(anchor, picIndex);
//                                            } else {
//                                                sxssCell.setCellValue(new String((byte[]) value));
//                                            }
//                                        } else {
//                                            String cellValue = xssCellValue.replace(excelFieldSrcKeyword, String.valueOf(field.get(dataObj)));
//                                            sxssCell.setCellValue(cellValue);
//                                        }
//                                    }
//                                }
//                            }
//                        }
//
//                        if (matchFlag == false) {
//                            SXSSFCell sxssCell = sxssrow.createCell(k, xssCell.getCellType());
//                            String value = null;
//                            if(xssCell.getCellType() == CellType.NUMERIC) {
//                            	value = String.valueOf(xssCell.getNumericCellValue());
//                            }else {
//                            	value = xssCell.getStringCellValue();
//                            }
//                            if (value != null && value.contains("${")) {
//                                String excelFieldSrcKeyword = value.substring(value.indexOf("${"), value.lastIndexOf("}") + 1);
//                                value = value.replace(excelFieldSrcKeyword, "");
//                            }
//                            CellStyle _sxssStyle = sxssfWorkbook.createCellStyle();
//                            if (callBackCellStyle != null) {
//                                _sxssStyle.cloneStyleFrom(xssCell.getCellStyle());
//                                sxssCell.setCellStyle(_sxssStyle);
//                                callBackCellStyle.callBack(sxssSheet, sxssCell, _sxssStyle);
//                            } else {
//                                _sxssStyle.cloneStyleFrom(xssCell.getCellStyle());
//                                sxssCell.setCellStyle(_sxssStyle);
//                            }
//                            sxssCell.setCellValue(value);
//                        }
//                    }
//                }
//            }
//        }
//
//        workbook.close();
//        ByteArrayOutputStream byteStream = new ByteArrayOutputStream();
//        sxssfWorkbook.write(byteStream);
//        byteStream.flush();
//        byteStream.close();
//        sxssfWorkbook.close();
//        sxssfWorkbook.dispose();
//        return byteStream.toByteArray();
//    }
//
//}
