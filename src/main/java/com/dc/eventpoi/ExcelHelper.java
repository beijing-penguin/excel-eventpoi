package com.dc.eventpoi;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dc.eventpoi.core.PoiUtils;
import com.dc.eventpoi.core.XlsxWriteStream;
import com.dc.eventpoi.core.entity.ExcelCell;
import com.dc.eventpoi.core.entity.ExcelRow;
import com.dc.eventpoi.core.entity.ListAndTableEntity;
import com.dc.eventpoi.core.enums.FileType;
import com.dc.eventpoi.core.func.DivideFunction;
import com.dc.eventpoi.core.func.DivideRoundFunction;
import com.dc.eventpoi.core.func.DivideTruncateFunction;
import com.dc.eventpoi.core.inter.CellStyleCallBack;
import com.dc.eventpoi.core.inter.ExcelEventStream;
import com.dc.eventpoi.core.inter.ExcelFunction;
import com.dc.eventpoi.core.inter.RowCallBack;
import com.dc.eventpoi.core.inter.SheetCallBack;

/**
 * excel操作
 *
 * @author beijing-penguin
 */
public class ExcelHelper {

	private static List<ExcelFunction> funcList = new ArrayList<>();
	static {
		funcList.add(new DivideFunction());
		funcList.add(new DivideRoundFunction());
		funcList.add(new DivideTruncateFunction());
	}
    /**
     * 导出表格 以及 列表数据
     * 
     * @param tempStream         模板文件流
     * @param listAndTableEntity 包含列表数据集合 和 表格数据对象
     * @param sheetIndex         sheetIndex
     * @param isClearPlaceholder 导出是否清楚占位符(默认true清除)
     * @param sheetCallBack      sheetCallBack
     * @param callBackCellStyle  callBackCellStyle
     * @return byte[]
     * @throws Exception Exception
     */
    public static byte[] exportExcel(InputStream tempStream, ListAndTableEntity listAndTableEntity, Integer sheetIndex, Boolean isClearPlaceholder, SheetCallBack sheetCallBack, CellStyleCallBack callBackCellStyle) throws Throwable {
        byte[] tempExcelBtye = PoiUtils.inputStreamToByte(tempStream);
        return exportExcel(tempExcelBtye, listAndTableEntity, sheetIndex, isClearPlaceholder, sheetCallBack, callBackCellStyle);
    }

    /**
     * 导出表格 以及 列表数据
     * 
     * @param tempExcelBtye      模板文件流
     * @param listAndTableEntity 包含列表数据集合 和 表格数据对象
     * @param sheetIndex         sheetIndex
     * @param isClearPlaceholder 导出是否清楚占位符(默认true清除)
     * @param sheetCallBack      sheetCallBack
     * @param callBackCellStyle  callBackCellStyle
     * @return byte[]
     * @throws Throwable 
     */
    public static byte[] exportExcel(byte[] tempExcelBtye, ListAndTableEntity listAndTableEntity, Integer sheetIndex, Boolean isClearPlaceholder, SheetCallBack sheetCallBack, CellStyleCallBack callBackCellStyle) throws Throwable {
    	XlsxWriteStream writeHelper = new XlsxWriteStream();
    	writeHelper.setAutoClearPlaceholder(true);
    	writeHelper.setSheetIndex(sheetIndex);
    	writeHelper.setAutoClearPlaceholder(isClearPlaceholder == null?true:isClearPlaceholder);
    	return writeHelper.exportExcel(tempExcelBtye, listAndTableEntity,funcList);
    }

    /**
     * 解析Excel为对象集合
     * 
     * @param excelTemplateStream   模版数据流
     * @param excelDataSourceStream Excel原数据流
     * @param clazz                 clazz
     * @param imageRead             是否支持图片格式读取（开启此功能，性能降低，内存消耗增加。）
     * @param <T>                   返回对象
     * @return 对象集合
     * @throws Exception IOException
     */
    public static <T> List<T> parseExcelToObject(InputStream excelTemplateStream, InputStream excelDataSourceStream, Class<T> clazz, boolean imageRead) throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024 * 4];
        int n = 0;
        while (-1 != (n = excelDataSourceStream.read(buffer))) {
            output.write(buffer, 0, n);
        }

        // 创建Workbook
        Workbook wb = null;
        // 创建sheet
        Sheet sheet = null;
        FileType fileType = PoiUtils.judgeFileType(new ByteArrayInputStream(output.toByteArray()));
        switch (fileType) {
        case XLS:
            wb = (HSSFWorkbook) WorkbookFactory.create(new ByteArrayInputStream(output.toByteArray()));
            break;
        case XLSX:
            wb = new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()));
            break;
        default:
            throw new Exception("filetype is unsupport");
        }
        // 获取excel sheet总数
        int sheetNumbers = wb.getNumberOfSheets();

        Map<String, byte[]> map = new HashMap<String, byte[]>();
        // 循环sheet
        for (int i = 0; i < sheetNumbers; i++) {

            sheet = wb.getSheetAt(i);

            switch (fileType) {
            case XLS:
                map.putAll(PoiUtils.getXlsPictures(i, (HSSFSheet) sheet));
                break;
            case XLSX:
                map.putAll(PoiUtils.getXlsxPictures(i, (XSSFSheet) sheet));
                break;
            default:
                throw new Exception("filetype is unsupport");
            }
        }
        wb.close();

        List<ExcelRow> dataList = ExcelHelper.parseExcelRowList(new ByteArrayInputStream(output.toByteArray()));
        List<ExcelRow> templeteList = ExcelHelper.parseExcelRowList(excelTemplateStream);
        checkTemplete(templeteList, dataList);

        if (map.size() > 0) {
            for (ExcelRow excelRow : dataList) {
                int rowIndex = excelRow.getRowIndex();
                int sheetIndex = excelRow.getSheetIndex();
                List<ExcelCell> cellList = excelRow.getCellList();
                for (Entry<String, byte[]> entry : map.entrySet()) {
                    int img_sheetIndex = Integer.parseInt(entry.getKey().split("-")[0]);
                    int img_rowIndex = Integer.parseInt(entry.getKey().split("-")[1]);
                    int img_cellIndex = Integer.parseInt(entry.getKey().split("-")[2]);
                    if (rowIndex == img_rowIndex && img_sheetIndex == sheetIndex) {
                        ExcelCell imgCell = new ExcelCell((short) img_sheetIndex, entry.getValue());
                        cellList.add(img_cellIndex, imgCell);
                        break;
                    }
                }
            }
        }
        return ExcelHelper.parseExcelToObject(templeteList, dataList, clazz);
    }

    public static <T> List<T> parseExcelToObject(InputStream excelTemplateStream, InputStream excelDataSourceStream, Class<T> clazz) throws Exception {
        List<ExcelRow> dataList = ExcelHelper.parseExcelRowList(excelDataSourceStream);
        List<ExcelRow> templeteList = ExcelHelper.parseExcelRowList(excelTemplateStream);
        checkTemplete(templeteList, dataList);
        return ExcelHelper.parseExcelToObject(templeteList, dataList, clazz);
    }

    /**
     * @param fileList     数据文件
     * @param templeteList 模板文件
     * @param clazz        类对象
     * @param <T>          T
     * @return 集合
     * @throws Exception IOException
     * @author beijing-penguin
     */
    public static <T> List<T> parseExcelToObject(List<ExcelRow> templeteList, List<ExcelRow> fileList, Class<T> clazz) throws Exception {
        List<T> rtn = new ArrayList<T>();
        List<ExcelCell> tempFieldList = new ArrayList<ExcelCell>();
        int size = fileList.size();
        int x = 0;
        int startRow = 0;
        for (int i = 0; i < templeteList.size(); i++) {
            if (templeteList.get(i).getCellList().get(0).getValue().startsWith("$")) {
                startRow = templeteList.get(i).getRowIndex();
                short sheetIndex = templeteList.get(i).getSheetIndex();
                tempFieldList = templeteList.get(i).getCellList();

                for (int j = (x + startRow); j < size; j++) {
                    ExcelRow row = fileList.get(j);
                    int rowIndex = row.getRowIndex();
                    if (rowIndex >= startRow && row.getSheetIndex() == sheetIndex) {
                        x++;
                        T obj = clazz.getDeclaredConstructor().newInstance();
                        List<ExcelCell> fieldList = row.getCellList();
                        for (ExcelCell fieldCell : fieldList) {
                            for (ExcelCell tempCell : tempFieldList) {
                                if (fieldCell.getIndex() == tempCell.getIndex()) {
                                    for (Field field : FieldUtils.getAllFields(clazz)) {
                                        if (!Modifier.isStatic(field.getModifiers())) {
                                            if (tempCell.getValue().contains(field.getName())) {
                                                field.setAccessible(true);
                                                if (fieldCell.getImgBytes() != null) {
                                                    // Object vall = getValueByFieldType(fieldCell.getImgBytes(),
                                                    // field.getType());
                                                    field.set(obj, fieldCell.getImgBytes());
                                                } else {
                                                    Object vall = PoiUtils.getValueByFieldType(fieldCell.getValue(), field.getType());
                                                    field.set(obj, vall);
                                                }
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        rtn.add(obj);
                    }
                }
            }
        }

        return rtn;
    }

    /**
     * 读取所有sheet数据
     *
     * @param baytes 文件
     * @return List
     * @throws Exception IOException
     * @author dc
     */
    public static List<ExcelRow> parseExcelRowList(byte[] baytes) throws Exception {
        return parseExcelRowList(new ByteArrayInputStream(baytes));
    }

    /**
     * 读取excel指定sheet数据
     *
     * @param baytes     文件数据
     * @param sheetIndex sheet工作簿索引号
     * @return List
     * @throws Exception IOException
     * @author dc
     */
    public static List<ExcelRow> parseExcelRowList(byte[] baytes, Integer sheetIndex) throws Exception {
        return parseExcelRowList(new ByteArrayInputStream(baytes), sheetIndex);
    }

    /**
     * 读取excel指定sheet数据
     *
     * @param file       文件
     * @param sheetIndex sheet工作簿索引号
     * @return List
     * @throws Exception IOException
     * @author dc
     */
    public static List<ExcelRow> parseExcelRowList(File file, Integer sheetIndex) throws Exception {
        return parseExcelRowList(new FileInputStream(file), sheetIndex);
    }

    /**
     * 读取所有sheet数据
     *
     * @param file 文件
     * @return List
     * @throws Exception IOException
     * @author dc
     */
    public static List<ExcelRow> parseExcelRowList(File file) throws Exception {
        return parseExcelRowList(new FileInputStream(file), null);
    }

    /**
     * 读取指定sheet数据
     *
     * @param inputSrc   excel源文件input输入流
     * @param sheetIndex sheet工作簿索引号
     * @return List
     * @throws Exception IOException
     * @author dc
     */
    public static List<ExcelRow> parseExcelRowList(InputStream inputSrc, Integer sheetIndex) throws Exception {
        List<ExcelRow> fileList = new ArrayList<ExcelRow>();
        ExcelEventStream fileStream = null;
        try {
            fileStream = ExcelEventStream.readExcel(inputSrc);
            fileStream.sheetAt(sheetIndex).rowStream(new RowCallBack() {
                @Override
                public void getRow(ExcelRow row) {
                    fileList.add(row);
                }
            });
        } catch (Exception e) {
            throw e;
        } finally {
            if (fileStream != null) {
                fileStream.close();
            }
        }
        return fileList;
    }

    /**
     * @param inputSrc excel源文件input输入流
     * @return List
     * @throws Exception IOException
     * @author dc
     */
    public static List<ExcelRow> parseExcelRowList(InputStream inputSrc) throws Exception {
        return parseExcelRowList(inputSrc, null);
    }

    /**
     * 模板与数据文件检查
     *
     * @param templeteList 模板文件
     * @param fileList     原始上传文件
     * @throws Exception IOException
     * @author beijing-penguin
     */
    public static void checkTemplete(List<ExcelRow> templeteList, List<ExcelRow> fileList) throws Exception {
        for (int i = 0; i < templeteList.size(); i++) {
            ExcelRow row = templeteList.get(i);
            List<ExcelCell> excelCell = row.getCellList();
            if (!excelCell.get(0).getValue().startsWith("${")) {
                for (int j = 0; j < excelCell.size(); j++) {
                    String tempValue = excelCell.get(j).getValue();
                    if (tempValue != null && !tempValue.startsWith("${")) {
                        String fileValue = fileList.get(i).getCellList().get(j).getValue();
                        if (!tempValue.equals(fileValue)) {
                            throw new Exception("fileList is not the same as templeteList[读取文件的excel头信息【" + fileValue + "】和模板头信息【" + tempValue + "】不匹配，文件格式不一致]");
                        }
                    }
                }
            } else {
                break;
            }
        }
    }
}
