/**
 * EventExcelHelper.java
 */
package com.dc.eventpoi;

import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.alibaba.fastjson.JSON;

/**
 * @Description: excel操作
 * @author 段超
 * @date: 2019年1月28日
 */
public class ExcelHelper {
    /**
     * 
     */
    private static Log LOG   = LogFactory.getLog(ExcelHelper.class);
    /**
     * 
     * @param fileList 数据文件
     * @param templeteList 模板文件
     * @param clazz 类对象
     * @return 集合
     * @throws Exception
     * @param  <T> T
     * @author 段超
     * @date 2019-01-28 14:25:20
     */
    public static <T> List<T> parseExcelToObject(List<ExcelRow> fileList,List<ExcelRow> templeteList, Class<T> clazz) throws Exception{
        int startRow = 0;
        List<ExcelCell> tempFieldList = new ArrayList<ExcelCell>();
        for (int i = 0; i < templeteList.size(); i++) {
            if(templeteList.get(i).getCellList().get(0).getValue().startsWith("$")) {
                startRow = templeteList.get(i).getRowIndex();
                tempFieldList = templeteList.get(i).getCellList();
                break;
            }
        }

        List<T> rtn = new ArrayList<T>();
        for (ExcelRow row : fileList) {
            int rowIndex = row.getRowIndex();
            if(rowIndex>=startRow) {
                T obj = clazz.newInstance();
                List<ExcelCell> fieldList = row.getCellList();
                for (ExcelCell fieldCell : fieldList) {
                    for (ExcelCell tempCell : tempFieldList) {
                        if(fieldCell.getIndex()==tempCell.getIndex()) {
                            for (Field field : FieldUtils.getAllFields(clazz)) {
                                if (!Modifier.isStatic(field.getModifiers())) {
                                    if (tempCell.getValue().contains(field.getName())) {
                                        field.setAccessible(true);
                                        Object vall = getValueByFieldType(fieldCell.getValue(), field.getType());
                                        field.set(obj, vall);
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
        return rtn;
    }
    /**
     * 读取所有sheet数据
     * @param baytes    文件
     * @return List
     * @throws Exception
     * @author dc
     */
    public static List<ExcelRow> parseExcelRowList(byte[] baytes) throws Exception{
        return parseExcelRowList(new ByteArrayInputStream(baytes));
    }
    /**
     * 读取excel指定sheet数据
     * @param baytes    文件数据
     * @param sheetIndex  sheet工作簿索引号
     * @return List
     * @throws Exception
     * @author dc
     */
    public static List<ExcelRow> parseExcelRowList(byte[] baytes,Integer sheetIndex) throws Exception{
        return parseExcelRowList(new ByteArrayInputStream(baytes),sheetIndex);
    }
    /**
     * 读取excel指定sheet数据
     * @param file    文件
     * @param sheetIndex  sheet工作簿索引号
     * @return List
     * @throws Exception
     * @author dc
     */
    public static List<ExcelRow> parseExcelRowList(File file,Integer sheetIndex) throws Exception{
        return parseExcelRowList(new FileInputStream(file),sheetIndex);
    }
    /**
     * 读取所有sheet数据
     * @param file    文件
     * @return List
     * @throws Exception
     * @author dc
     */
    public static List<ExcelRow> parseExcelRowList(File file) throws Exception{
        return parseExcelRowList(new FileInputStream(file),null);
    }
    /**
     * 读取指定sheet数据
     * @param inputSrc    excel源文件input输入流
     * @param sheetIndex  sheet工作簿索引号
     * @return List
     * @throws Exception
     * @author dc
     */
    public static List<ExcelRow> parseExcelRowList(InputStream inputSrc,Integer sheetIndex) throws Exception{
        List<ExcelRow> fileList  =new ArrayList<ExcelRow>();
        ExcelEventStream fileStream = null;
        try {
            fileStream = ExcelEventStream.readExcel(inputSrc);
            fileStream.sheetAt(sheetIndex).rowStream(new RowCallBack() {
                @Override
                public void getRow(ExcelRow row) {
                    fileList.add(row);
                }
            });
        }catch (Exception e) {
            throw e;
        }finally {
            if(fileStream!=null) {
                fileStream.close();
            }
        }
        return fileList;
    }
    /**
     * @param inputSrc    excel源文件input输入流
     * @return List
     * @throws Exception
     * @author dc
     */
    public static List<ExcelRow> parseExcelRowList(InputStream inputSrc) throws Exception{
        return parseExcelRowList(inputSrc,null);
    }
    /**
     * 模板与数据文件检查
     * @param fileList 原始上传文件
     * @param templeteList 模板文件
     * @throws Exception
     * @author 段超
     * @date 2019-01-28 14:24:09
     */
    public static void checkTemplete(List<ExcelRow> fileList,List<ExcelRow> templeteList) throws Exception{
        for (int i = 0; i < templeteList.size(); i++) {
            ExcelRow row = templeteList.get(i);
            List<ExcelCell> excelCell = row.getCellList();
            if(!excelCell.get(0).getValue().startsWith("${")) {
                if(!JSON.toJSONString(templeteList.get(i)).equals(JSON.toJSONString(fileList.get(i)))) {
                    throw new Exception("fileList is not the same as templeteList");
                }
            }else {
                break;
            }
        }
    }

    /**
     * @param value     任意数据类型对象
     * @param fieldType 转化后的类型
     * @return Object
     * @throws Exception
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
        }else if(v.trim().length()==0) {
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
        } else if (type.equals("Boolean") || type.equals("boolean")) {
            return Boolean.parseBoolean(v);
        } else if (type.equals("BigDecimal")) {
            return new BigDecimal(v);
        } else if (type.equals("BigInteger")) {
            return new BigInteger(v);
        } else if (type.equals("Date")) {
            SimpleDateFormat sdf = new SimpleDateFormat(getDateFormat(v));
            //不允许底层java自动日期进行计算，直接抛出异常
            sdf.setLenient(false);
            Date date = sdf.parse(v);
            return date;
        }
        throw new Exception(type + " is unsupported");
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
     * 删除模板中的固定格式
     * @param inputSrc 源模板文件
     * @param sheetIndex 工作簿索引下标
     * @return ByteArrayOutputStream
     * @throws Exception
     * @author 段超
     * @date 2019-01-30 16:59:13
     */
    public static ByteArrayOutputStream deleteTempleteFormat(InputStream inputSrc,int sheetIndex) throws Exception {
        Workbook workbook = WorkbookFactory.create(inputSrc);
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        int totalRow = sheet.getPhysicalNumberOfRows();
        for (int i = sheet.getFirstRowNum(); i < totalRow; i++) {
            Row row = sheet.getRow(i); 
            if(row!=null) {
                for (int j = row.getFirstCellNum(),totalCell=row.getPhysicalNumberOfCells(); j < totalCell; j++) {
                    Cell cell = row.getCell(j);
                    if(cell!=null) {
                        //cell.setCellType(CellType.STRING);
                        String value = cell.getStringCellValue();
                        if(value!=null && value.startsWith("${")) {
                            sheet.removeRow(row);
                            sheet.shiftRows(i+1, i+1+1, -1);
                            break;
                        }
                    }
                }
            }
        }
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        workbook.write(os);
        try {
            os.flush();
        }catch (Exception e) {
            LOG.error("", e);
        }
        try {
            inputSrc.close();
        }catch (Exception e) {
            LOG.error("", e);
        }
        try {
            os.close();
        }catch (Exception e) {
            LOG.error("", e);
        }
        return os;
    }
    /**
     * 导出zip文件,或者导出xlxs文件
     * @param templeteFileName 模板文件名
     * @param templete 模板文件数据
     * @param dataList 对象数据集合
     * @param sheetIndex 工作簿
     * @param isZip 是否压缩成zip文件
     * @return byte[]
     * @throws Exception
     * @author 段超
     * @date 2019-02-22 14:29:52
     */
    public static byte[] exportExcel(String templeteFileName,byte[] templete,List<?> dataList,int sheetIndex,boolean isZip) throws Exception {
        FileType fileType = judgeFileType(new ByteArrayInputStream(templete));
        if(templeteFileName==null) {
            if(fileType==FileType.XLSX) {
                templeteFileName = "file.xlxs";
            }else {
                templeteFileName = "file.xls";
            }
        }
        Integer startRow = null;
        int dataTotal = 45000;
        if(isZip) {
            if(fileType==FileType.XLSX) {
                dataTotal = 30000;//3w
            }else {
                dataTotal = 30000;//3w
            }
        }else {
            dataTotal = Integer.MAX_VALUE;
        }

        List<ExcelRow> rowList  =new ArrayList<ExcelRow>();
        List<ExcelCell> keyCellList = null;
        ExcelEventStream fileStream = ExcelEventStream.readExcel(templete);
        fileStream.sheetAt(sheetIndex).rowStream(new RowCallBack() {
            @Override
            public void getRow(ExcelRow row) {
                rowList.add(row);
            }
        });
        for (int i = 0; i < rowList.size(); i++) {
            if(startRow!=null) {
                break;
            }
            ExcelRow row = rowList.get(i);
            List<ExcelCell> cellList = row.getCellList();
            for (int j = 0; j < cellList.size(); j++) {
                ExcelCell cell = cellList.get(j);
                if(cell.getValue().startsWith("${")) {
                    startRow = row.getRowIndex();
                    keyCellList = cellList;
                    break;
                }
            }
        }
        int len = dataList.size();
        int l = len/dataTotal+(len%dataTotal!=0?1:0);
        Map<String, byte[]> fileDataMap = new LinkedHashMap<String, byte[]>();
        for (int i = 0; i < l; i++) {
            Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(templete));
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setAlignment(HorizontalAlignment.CENTER); // 水平居中
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 上下居中
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);

            for (int j = i*dataTotal; j < (i+1)*dataTotal; j++) {
                if(j>=len) {
                    break;
                }
                Row row = sheet.createRow(j+startRow-i*dataTotal);
                Object obj = dataList.get(j);
                for (int k = 0; k < keyCellList.size(); k++) {
                    ExcelCell cellField = keyCellList.get(k);
                    String excelField = cellField.getValue().substring(cellField.getValue().indexOf("${") + 2, cellField.getValue().lastIndexOf("}"));
                    Field[] fieldArr = FieldUtils.getAllFields(obj.getClass());
                    for (Field field : fieldArr) {
                        if (!Modifier.isStatic(field.getModifiers()) && field.getName().equals(excelField)) {
                            Cell cell = row.createCell(cellField.getIndex(), CellType.STRING);
                            field.setAccessible(true);
                            Object value = field.get(obj);
                            if (value != null && value.toString().trim().length() > 0) {
                                cell.setCellValue(String.valueOf(value));
                                cell.setCellStyle(cellStyle);
                            }
                        }
                    }
                }
            }
            ByteArrayOutputStream byteStream =new ByteArrayOutputStream();
            workbook.write(byteStream);
            try {
                byteStream.flush();
            }catch (Exception e) {
                LOG.error("",e);
                throw e;
            }
            try {
                byteStream.close();
            }catch (Exception e) {
                LOG.error("",e);
                throw e;
            }
            workbook.close();
            workbook = null;
            String newName = "";
            if(templeteFileName.lastIndexOf(".")!=-1) {
                newName = templeteFileName.substring(0, templeteFileName.lastIndexOf("."))+"_"+i+templeteFileName.substring(templeteFileName.lastIndexOf("."));
            }else {
                newName = templeteFileName+"_"+i;
            }
            fileDataMap.put(newName, byteStream.toByteArray());
        }
        if(isZip) {
            throw new Exception("zip is unsupport");
            //return ZipUtils.batchCompress(fileDataMap);
        }else {
            return fileDataMap.values().iterator().next();
        }
    }


    /**
     * 导出zip文件,或者导出xlxs文件
     * @param templeteFileName 模板文件名
     * @param templete 模板文件数据
     * @param dataList 对象数据集合
     * @param sheetIndex 工作簿
     * @param isZip 是否压缩成zip文件
     * @return byte[]
     * @throws Exception
     * @author 段超
     * @date 2019-02-22 14:29:52
     */
    public static byte[] exportTitleExcel(byte[] templete,Object data,int sheetIndex) throws Exception {
        List<ExcelRow> rowList  = new ArrayList<ExcelRow>();
        List<ExcelCell> keyCellList = new ArrayList<ExcelCell>();
        ExcelEventStream fileStream = ExcelEventStream.readExcel(templete);
        fileStream.sheetAt(sheetIndex).rowStream(new RowCallBack() {
            @Override
            public void getRow(ExcelRow row) {
                rowList.add(row);
            }
        });
        for (int i = 0; i < rowList.size(); i++) {
            ExcelRow row = rowList.get(i);
            List<ExcelCell> cellList = row.getCellList();
            for (int j = 0; j < cellList.size(); j++) {
                ExcelCell cell = cellList.get(j);
                if(cell.getValue().startsWith("${")) {
                    keyCellList.addAll(cellList);
                    break;
                }
            }
        }
        Map<String, byte[]> fileDataMap = new LinkedHashMap<String, byte[]>();
        Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(templete));
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER); // 水平居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 上下居中
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);

        for (int j = 0; j < sheet.getLastRowNum(); j++) {
            Row row = sheet.getRow(j);
            for (int cellnum=0;cellnum<=row.getLastCellNum();cellnum++){
                Cell cell = row.getCell(cellnum);
                if (cell != null ) {
                    String cellValue = cell.getStringCellValue();
                    if(cellValue!=null && cellValue.startsWith("${")) {
                        String excelField = cellValue.substring(cellValue.indexOf("${") + 2, cellValue.lastIndexOf("}"));
                        Field[] fieldArr = FieldUtils.getAllFields(data.getClass());
                        for (Field field : fieldArr) {
                            if (!Modifier.isStatic(field.getModifiers()) && field.getName().equals(excelField)) {
                                field.setAccessible(true);
                                Object value = field.get(data);
                                if (value != null && value.toString().trim().length() > 0) {
                                    cell.setCellValue(String.valueOf(value));
                                    cell.setCellStyle(cellStyle);
                                }
                            }
                        }
                    }
                }
            }
        }
        ByteArrayOutputStream byteStream =new ByteArrayOutputStream();
        workbook.write(byteStream);
        try {
            byteStream.flush();
        }catch (Exception e) {
            LOG.error("",e);
            throw e;
        }
        try {
            byteStream.close();
        }catch (Exception e) {
            LOG.error("",e);
            throw e;
        }
        workbook.close();
        workbook = null;
        String newName = "temp";
        fileDataMap.put(newName, byteStream.toByteArray());
        return fileDataMap.values().iterator().next();
    }



    /**
     * 导出zip数据
     * @param templete 模板数据
     * @param dataList 对象数据集合
     * @return byte[]
     * @throws Exception
     * @author 段超
     * @date 2019-02-22 14:33:33
     */
    public static byte[] exportExcel(byte[] templete,List<?> dataList) throws Exception {
        return exportExcel(null,templete, dataList,0,true);
    }
    /**
     * 导出zip数据
     * @param templeteFileName 模板文件名
     * @param templete templete
     * @param dataList dataList
     * @return byte[]
     * @throws Exception
     * @author 段超
     * @date 2019-02-22 14:33:33
     */
    public static byte[] exportExcel(String templeteFileName , byte[] templete,List<?> dataList) throws Exception {
        return exportExcel(templeteFileName,templete, dataList,0,true);
    }
    /**
     * 导出zip数据
     * @param templete 模板文件流
     * @param dataList 对象数据集合
     * @param isZip 是否压缩
     * @return byte[]
     * @throws Exception
     * @author 段超
     * @date 2019-02-22 14:33:33
     */
    public static byte[] exportExcel(byte[] templete,List<?> dataList,boolean isZip) throws Exception {
        return exportExcel(null,templete, dataList,0,isZip);
    }
    /**
     * 流转byte[]
     * @param is is
     * @return byte[]
     * @throws Exception
     * @author 段超
     * @date 2019-02-22 15:16:47
     */
    public static byte[] inputStreamToByte(InputStream is) throws Exception{
        BufferedInputStream bis = new BufferedInputStream(is);
        byte [] a = new byte[1000];
        int len = 0;
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        while((len = bis.read(a))!=-1){
            bos.write(a, 0, len);
        }
        bis.close();
        bos.close();
        return bos.toByteArray();
    }

    /**
     * 判断文件类型
     * @param inp 数据流
     * @return FileType
     * @throws Exception
     * @author 段超
     * @date 2019-02-25 11:27:10
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
    /**
     * 获取cell值
     * @param cellList cell集合
     * @param cellIndex 索引号
     * @param returnClass 返回类型
     * @param <T> 返回类型
     * @return T
     * @throws Exception
     * @author 段超
     * @date 2019-02-26 18:36:22
     */
    @SuppressWarnings("unchecked")
    public static <T> T getValueBy(List<ExcelCell> cellList ,int cellIndex,Class<? extends T> returnClass) throws Exception {
        for (int i = 0; i < cellList.size(); i++) {
            ExcelCell cell = cellList.get(i);
            if(cell.getIndex()==cellIndex) {
                return (T) getValueByFieldType(cell.getValue(), returnClass);
            }
        }
        return null;
    }
    /**
     * 获取cell值
     * @param cellList cell集合
     * @param cellIndex 索引号
     * @return String
     * @throws Exception
     * @author 段超
     * @date 2019-02-26 18:46:16
     */
    public static String getValueBy(List<ExcelCell> cellList ,int cellIndex) throws Exception {
        return getValueBy(cellList, cellIndex,String.class);
    }
    /**
     * 获取值
     * @param rowList 行集合
     * @param rowIndex 行下标
     * @param cellIndex 列下标
     * @param returnClass 返回值类型
     * @param <T> 返回类型
     * @return T
     * @throws Exception
     * @author 段超
     * @date 2019-02-26 19:19:49
     */
    @SuppressWarnings("unchecked")
    public static <T> T getValueBy(List<ExcelRow> rowList ,int rowIndex,int cellIndex,Class<? extends T> returnClass) throws Exception {
        for (int i = 0; i < rowList.size(); i++) {
            ExcelRow row = rowList.get(i);
            if(row.getRowIndex()>rowIndex) {
                break;
            }
            if(row.getRowIndex()==rowIndex) {
                List<ExcelCell> cellList = row.getCellList();
                for (int j = 0; j < cellList.size(); j++) {
                    ExcelCell cell = cellList.get(j);
                    if(cell.getIndex()==cellIndex) {
                        return (T) getValueByFieldType(cell.getValue(), returnClass);
                    }
                }
            }
        }
        return null;
    }
    /**
     * 获取值
     * @param rowList 行集合
     * @param rowIndex 行下标
     * @param cellIndex 列下标
     * @return String
     * @throws Exception
     * @author 段超
     * @date 2019-02-26 19:19:49
     */
    public static String getValueBy(List<ExcelRow> rowList ,int rowIndex,int cellIndex) throws Exception {
        return getValueBy(rowList, rowIndex,cellIndex,String.class);
    }
}
