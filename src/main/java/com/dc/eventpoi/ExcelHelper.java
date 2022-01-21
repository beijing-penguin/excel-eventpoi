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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFDrawing;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dc.eventpoi.core.PoiUtils;
import com.dc.eventpoi.core.entity.ExcelCell;
import com.dc.eventpoi.core.entity.ExcelRow;
import com.dc.eventpoi.core.entity.ExportExcelCell;
import com.dc.eventpoi.core.entity.ListAndTableEntity;
import com.dc.eventpoi.core.enums.FileType;
import com.dc.eventpoi.core.inter.CellStyleCallBack;
import com.dc.eventpoi.core.inter.ExcelEventStream;
import com.dc.eventpoi.core.inter.RowCallBack;
import com.dc.eventpoi.core.inter.SheetCallBack;

/**
 * excel操作
 *
 * @author beijing-penguin
 */
public class ExcelHelper {

    /**
     * 导出表格 以及 列表数据
     * 
     * @param tempStream         模板文件流
     * @param listAndTableEntity 包含列表数据集合 和 表格数据对象
     * @param sheetIndex         sheetIndex
     * @param isClearPlaceholder 是否清除占位符
     * @param sheetCallBack      sheetCallBack
     * @param callBackCellStyle  callBackCellStyle
     * @return byte[]
     * @throws Exception Exception
     */
    public static byte[] exportExcel(InputStream tempStream, ListAndTableEntity listAndTableEntity, Integer sheetIndex, Boolean isClearPlaceholder, SheetCallBack sheetCallBack, CellStyleCallBack callBackCellStyle) throws Exception {
        byte[] tempExcelBtye = PoiUtils.inputStreamToByte(tempStream);
        return exportExcel(tempExcelBtye, listAndTableEntity, sheetIndex, isClearPlaceholder, sheetCallBack, callBackCellStyle);
    }

    /**
     * 导出表格 以及 列表数据
     * 
     * @param tempExcelBtye      模板文件流
     * @param listAndTableEntity 包含列表数据集合 和 表格数据对象
     * @param sheetIndex         sheetIndex
     * @param isClearPlaceholder 是否清楚占位符
     * @param sheetCallBack      sheetCallBack
     * @param callBackCellStyle  callBackCellStyle
     * @return byte[]
     * @throws Exception Exception
     */
    public static byte[] exportExcel(byte[] tempExcelBtye, ListAndTableEntity listAndTableEntity, Integer sheetIndex, Boolean isClearPlaceholder, SheetCallBack sheetCallBack, CellStyleCallBack callBackCellStyle) throws Exception {
        if (isClearPlaceholder == null) {
            isClearPlaceholder = true;
        }

        Workbook workbook_import = null;
        FileType fileType = PoiUtils.judgeFileType(new ByteArrayInputStream(tempExcelBtye));
        if (fileType == FileType.XLSX) {
            workbook_import = new XSSFWorkbook(new ByteArrayInputStream(tempExcelBtye));
        } else {
            workbook_import = (HSSFWorkbook) WorkbookFactory.create(new ByteArrayInputStream(tempExcelBtye));
        }

        SXSSFWorkbook sxssfWorkbook_export = new SXSSFWorkbook();

        int sheetStart = 0;
        int sheetEnd = workbook_import.getNumberOfSheets();
        if (sheetIndex != null) {
            sheetStart = sheetIndex;
            sheetEnd = sheetIndex + 1;
        }
        for (int i = sheetStart; i < sheetEnd; i++) {
            SXSSFSheet sxssSheet_export = sxssfWorkbook_export.createSheet(workbook_import.getSheetName(i));
            if (sheetCallBack != null) {
                sheetCallBack.callBack(sxssSheet_export);
            }

            SXSSFDrawing patriarch = (SXSSFDrawing) sxssSheet_export.createDrawingPatriarch();
            Sheet xsssheet_import = workbook_import.getSheetAt(i);
            int sheetMergerCount = xsssheet_import.getNumMergedRegions();

            int rowNum = xsssheet_import.getPhysicalNumberOfRows();
            int offset = 0;
            int listCount = 0;
            for (int j = 0; j < rowNum; j++) {
                for (int ii = 0; ii < sheetMergerCount; ii++) {
                    CellRangeAddress mergedRegionAt = xsssheet_import.getMergedRegion(ii);
                    if (mergedRegionAt.getFirstRow() == j) {
                        mergedRegionAt.setFirstRow(mergedRegionAt.getFirstRow() + offset - listCount);
                        mergedRegionAt.setLastRow(mergedRegionAt.getLastRow() + offset - listCount);
                        sxssSheet_export.addMergedRegion(mergedRegionAt);
                    }
                }

                Row xssrow_import = xsssheet_import.getRow(j);
                int xssCellNum_import = xssrow_import.getPhysicalNumberOfCells();
                boolean breakFlag = false;

                SXSSFRow sxssrow_export = sxssSheet_export.createRow(j + offset - listCount);
                sxssrow_export.setHeight(xssrow_import.getHeight());

                for (int k_import = 0; k_import < xssCellNum_import; k_import++) {
                    if (breakFlag) {
                        break;
                    }
                    Cell xssCell_import = xssrow_import.getCell(k_import);
                    sxssSheet_export.setColumnWidth(k_import, xsssheet_import.getColumnWidth(k_import));
                    if (xssCell_import == null) {
                    } else {
                        boolean matchFlag = false;
                        String xssCellValue_import = PoiUtils.getCellValue(xssCell_import);
                        
                        if (xssCellValue_import != null && xssCellValue_import.contains("${")) {
                            String keyName_import = xssCellValue_import.substring(xssCellValue_import.indexOf("${") + 2, xssCellValue_import.lastIndexOf("}"));
                            String excelFieldSrcKeyword_import = xssCellValue_import.substring(xssCellValue_import.indexOf("${"), xssCellValue_import.lastIndexOf("}") + 1);

                            if (matchFlag) {
                                break;
                            }
                            List<?> dataList = (List<?>) listAndTableEntity.getDataList();
                            if (dataList != null && dataList.size() > 0) {
                                Object tempData = dataList.get(0);
                                if (FieldUtils.getField(tempData.getClass(), keyName_import, true) == null) {
                                } else {
                                    List<ExportExcelCell> keyCellList = new ArrayList<ExportExcelCell>();
                                    for (int kk_import = k_import; kk_import < xssCellNum_import; kk_import++) {
                                        Cell xssCell_kk_import = xssrow_import.getCell(kk_import);
                                        CellType type = xssCell_kk_import.getCellType();
                                        CellStyle _sxssStyle = sxssfWorkbook_export.createCellStyle();
                                        _sxssStyle.cloneStyleFrom(xssCell_kk_import.getCellStyle());

                                        ExportExcelCell ee = new ExportExcelCell((short) xssCell_kk_import.getColumnIndex(), xssCell_kk_import.getStringCellValue(), _sxssStyle);
                                        ee.setCellType(type);
                                        keyCellList.add(ee);
                                    }
                                    breakFlag = true;
                                    matchFlag = true;
                                    listCount++;
                                    for (int y = 0, len = dataList.size(); y < len; y++) {
                                        final int create_row_num = j + offset;
                                        offset++;

                                        Object srcData = dataList.get(y);
                                        SXSSFRow sxssrow_export_2 = sxssSheet_export.createRow(create_row_num);
                                        sxssrow_export_2.setHeight(xssrow_import.getHeight());
                                        // 判断是否存在当前行，如果模板中，存在当前行，则复制当前行前几列的模样
                                        if (y == 0) {
                                            for (int cell_index_import = k_import - 1; cell_index_import >= 0; cell_index_import--) {
                                                Cell beforCell_import = xssrow_import.getCell(cell_index_import);
                                                if (beforCell_import != null) {
                                                    SXSSFCell sxssCell_export = sxssrow_export_2.createCell(cell_index_import, beforCell_import.getCellType());
                                                    sxssCell_export.setCellStyle(beforCell_import.getCellStyle());

                                                    String setvv = PoiUtils.getCellValue(beforCell_import);
                                                    if (setvv == null) {
                                                        setvv = "";
                                                    }
                                                    sxssCell_export.setCellValue(setvv);
                                                }
                                            }
                                        }
                                        for (int x = k_import; x < xssCellNum_import; x++) {

                                            ExportExcelCell curCell_import = null;
                                            String vv = null;
                                            for (ExportExcelCell exportCell : keyCellList) {
                                                if (exportCell.getIndex() == x) {
                                                    curCell_import = exportCell;
                                                    vv = exportCell.getValue();
                                                    break;
                                                }
                                            }
                                            // curCell.getCellStyle().setFillForegroundColor(IndexedColors.AQUA.getIndex());
                                            // curCell.getCellStyle().setFillPattern(FillPatternType.SOLID_FOREGROUND);
                                            String _keyName = null;
                                            Field field = null;
                                            String excelFieldSrcKeyword2 = null;
                                            if (vv != null && vv.contains("${")) {
                                                _keyName = vv.substring(vv.indexOf("${") + 2, vv.lastIndexOf("}"));
                                                field = FieldUtils.getField(srcData.getClass(), _keyName, true);
                                                excelFieldSrcKeyword2 = vv.substring(vv.indexOf("${"), vv.lastIndexOf("}") + 1);
                                            }

                                            if (field != null && field.get(srcData) != null) {
                                                SXSSFCell sxssCell_export = sxssrow_export_2.createCell(x, curCell_import.getCellType());
                                                if (callBackCellStyle != null) {
                                                    callBackCellStyle.callBack(sxssSheet_export, sxssCell_export, curCell_import.getCellStyle());
                                                    sxssCell_export.setCellStyle(curCell_import.getCellStyle());
                                                } else {
                                                    sxssCell_export.setCellStyle(curCell_import.getCellStyle());
                                                }

                                                Object value = field.get(srcData);
                                                if (value instanceof byte[]) {
                                                    if (PoiUtils.getImageType((byte[]) value) != null) {
                                                        XSSFClientAnchor anchor_export = new XSSFClientAnchor(0, 0, 0, 0, x, sxssrow_export_2.getRowNum(), x + 1, sxssrow_export_2.getRowNum() + 1);
                                                        int picIndex = sxssfWorkbook_export.addPicture((byte[]) value, HSSFWorkbook.PICTURE_TYPE_JPEG);
                                                        patriarch.createPicture(anchor_export, picIndex);
                                                    } else {
                                                        sxssCell_export.setCellValue(new String((byte[]) value));
                                                    }
                                                } else {
                                                    sxssCell_export.setCellValue(String.valueOf(value));
                                                }
                                            } else {
                                                SXSSFCell sxssCell_export = sxssrow_export_2.createCell(x, curCell_import.getCellType());
                                                if (callBackCellStyle != null) {
                                                    callBackCellStyle.callBack(sxssSheet_export, sxssCell_export, curCell_import.getCellStyle());
                                                    sxssCell_export.setCellStyle(curCell_import.getCellStyle());
                                                } else {
                                                    sxssCell_export.setCellStyle(curCell_import.getCellStyle());
                                                }
                                                if (vv == null) {
                                                    vv = "";
                                                }
                                                String cellValue = vv;
                                                if (excelFieldSrcKeyword2 != null) {
                                                    cellValue = cellValue.replace(excelFieldSrcKeyword2, "");
                                                }
                                                sxssCell_export.setCellValue(cellValue);
                                            }
                                        }
                                    }
                                }
                            }

                            if (matchFlag == false) {
                                if (listAndTableEntity.getTableList() != null) {
                                    for (Object tableObject : listAndTableEntity.getTableList()) {
                                        Field field = FieldUtils.getField(tableObject.getClass(), keyName_import, true);
                                        if (field != null) {
                                            matchFlag = true;
                                            SXSSFCell sxssCell_export = sxssrow_export.createCell(k_import, xssCell_import.getCellType());
                                            CellStyle sxssStyle_export = sxssfWorkbook_export.createCellStyle();
                                            if (callBackCellStyle != null) {
                                                sxssStyle_export.cloneStyleFrom(xssCell_import.getCellStyle());
                                                sxssCell_export.setCellStyle(sxssStyle_export);
                                                callBackCellStyle.callBack(sxssSheet_export, sxssCell_export, sxssStyle_export);
                                            } else {
                                                sxssStyle_export.cloneStyleFrom(xssCell_import.getCellStyle());
                                                sxssCell_export.setCellStyle(sxssStyle_export);
                                            }

                                            Object value = field.get(tableObject);
                                            if (value instanceof byte[]) {
                                                if (PoiUtils.getImageType((byte[]) value) != null) {
                                                    XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, k_import, sxssrow_export.getRowNum(), k_import + 1, sxssrow_export.getRowNum() + 1);
                                                    int picIndex = sxssfWorkbook_export.addPicture((byte[]) value, HSSFWorkbook.PICTURE_TYPE_JPEG);
                                                    patriarch.createPicture(anchor, picIndex);
                                                } else {
                                                    sxssCell_export.setCellValue(new String((byte[]) value));
                                                }
                                            } else {
                                                String cellValue = xssCellValue_import.replace(excelFieldSrcKeyword_import, String.valueOf(value));
                                                sxssCell_export.setCellValue(cellValue);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (matchFlag == false) {// 单元格没匹配到，则清除单元格得占位符
                            SXSSFCell sxssCell_export = sxssrow_export.createCell(k_import, xssCell_import.getCellType());
                            String value = null;
                            if (xssCell_import.getCellType() == CellType.NUMERIC) {
                                value = String.valueOf(xssCell_import.getNumericCellValue());
                            } else {
                                value = xssCell_import.getStringCellValue();
                            }
                            if (value != null && value.contains("${") && isClearPlaceholder != null && isClearPlaceholder == true) {
                                String excelFieldSrcKeyword = value.substring(value.indexOf("${"), value.lastIndexOf("}") + 1);
                                value = value.replace(excelFieldSrcKeyword, "");
                            }
                            CellStyle sxssStyle_export = sxssfWorkbook_export.createCellStyle();
                            if (callBackCellStyle != null) {
                                sxssStyle_export.cloneStyleFrom(xssCell_import.getCellStyle());
                                sxssCell_export.setCellStyle(sxssStyle_export);
                                callBackCellStyle.callBack(sxssSheet_export, sxssCell_export, sxssStyle_export);
                            } else {
                                sxssStyle_export.cloneStyleFrom(xssCell_import.getCellStyle());
                                sxssCell_export.setCellStyle(sxssStyle_export);
                            }
                            sxssCell_export.setCellValue(value);
                        }
                    }
                }
            }
        }

        workbook_import.close();
        ByteArrayOutputStream byteStream = new ByteArrayOutputStream();
        sxssfWorkbook_export.write(byteStream);
        byteStream.flush();
        byteStream.close();
        sxssfWorkbook_export.close();
        sxssfWorkbook_export.dispose();
        return byteStream.toByteArray();
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
