package com.dc.eventpoi.core;

import com.dc.eventpoi.core.entity.ExcelCell;
import com.dc.eventpoi.core.entity.ExcelRow;
import com.dc.eventpoi.core.enums.CellDataTypeEnum;
import com.dc.eventpoi.core.inter.BaseCallBack;
import com.dc.eventpoi.core.inter.ExcelEventStream;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.model.SharedStrings;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.SAXParserFactory;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 * 新版xlxs事件流解析对象
 *
 * @author beijing-penguin
 */
public class ExcelXlsxStream extends BaseExcelStream implements ExcelEventStream {
    /**
     *
     */
    private OPCPackage pkg = null;
    /**
     *
     */
    private XSSFReader r = null;
    /**
     *
     */
    private SharedStrings sst = null;
    /**
     *
     */
    private XMLReader parser = null;
    /**
     *
     */
    private InputStream is = null;
    /**
     *
     */
    private SheetIterator sheets = null;
    /**
     *
     */
    private short sheetIndex = 0;
    /**
     *
     */
    private String sheetName;
    /**
     *
     */
    private BaseCallBack baseCallBack;
    /**
     *
     */
    private List<String> sheetList = new ArrayList<String>();
    /**
     *
     */
    private DefaultHandler defaultHandler = new DefaultHandler() {

        private String lastContents;
        //private boolean nextIsString;
        private CellDataTypeEnum cellDataType;
        //private boolean inlineStr;
        private String cellNo;
        private int curRowNum = 0;
        private List<ExcelCell> valueList = new ArrayList<ExcelCell>();

        //字符缓存优化
//		private final LruCache<Integer,String> lruCache = new LruCache<>(8000);
//		class LruCache<A,B> extends LinkedHashMap<A, B> {
//			private static final long serialVersionUID = 1L;
//			private final int maxEntries;
//			LruCache(final int maxEntries) {
//				super(maxEntries + 1, 1.0f, true);
//				this.maxEntries = maxEntries;
//			}
//
//			@Override
//			protected boolean removeEldestEntry(final Map.Entry<A, B> eldest) {
//				return super.size() > maxEntries;
//			}
//		}
        @Override
        public void characters(char[] ch, int start, int length) throws SAXException {
            lastContents += new String(ch, start, length);
        }

        // t="s" ,it's means String
        // t="str" ,it's means String,but does not need to be read in the 'sharedStrings.xml'
        // t="inlineStr" ,it's means String
        // t="b" ,it's means Boolean
        // t="e" ,it's means Error
        // t="n" ,it's means Number
        // t is null ,it's means Empty or Number
        @Override
        public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
            // c => cell
            if (name.equals("c")) {
                cellNo = attributes.getValue("r");
                //String cellType = attributes.getValue("t");
                cellDataType = CellDataTypeEnum.buildFromCellType(attributes.getValue("t"));
                //nextIsString = cellType != null && cellType.equals("s");
                //inlineStr = cellType != null && cellType.equals("inlineStr");
            }
            lastContents = "";
        }

        @Override
        public void endElement(String uri, String localName, String name) throws SAXException {
//			if(nextIsString) {
//				Integer idx = Integer.valueOf(lastContents);
//				//性能优化，数据量越大，提高约2秒的速度
//				lastContents = lruCache.get(idx);
//				if (lastContents == null && !lruCache.containsKey(idx)) {
//					lastContents = sst.getItemAt(idx).toString();
//					lruCache.put(idx, lastContents);
//				}
//				nextIsString = false;
//			}
            if (name.equals("v") || (cellDataType == CellDataTypeEnum.IN_LINE_STR && name.equals("c"))) {
                if (cellDataType == CellDataTypeEnum.EMPTY) {
                    //lastContents = sst.getItemAt(idx).toString();
                    lastContents = BigDecimal.valueOf(Double.parseDouble(lastContents)).stripTrailingZeros().toPlainString();
                } else {
                    Integer idx = Integer.valueOf(lastContents);
                    lastContents = sst.getItemAt(idx).toString();
                }


                String[] cellNoArr = parseCellNo(cellNo);
                short cellNum = (short) (excelColStrToNum(cellNoArr[0]) - 1);
                int rowNum = Integer.parseInt(cellNoArr[1]) - 1;
                if (curRowNum == rowNum) {
                    valueList.add(new ExcelCell(cellNum, lastContents));
                } else {
                    ExcelRow excelRow = new ExcelRow();
                    excelRow.setRowIndex(curRowNum);
                    excelRow.setCellList(valueList);
                    excelRow.setSheetIndex(sheetIndex);
                    excuteCallBack(baseCallBack, excelRow);
                    valueList = new ArrayList<ExcelCell>();
                    valueList.add(new ExcelCell(cellNum, lastContents));
                    curRowNum = rowNum;
                }
            }
        }

        @Override
        public void endDocument() throws SAXException {
            if (valueList.size() != 0) {
                ExcelRow excelRow = new ExcelRow();
                excelRow.setRowIndex(curRowNum);
                excelRow.setCellList(valueList);
                excelRow.setSheetIndex(sheetIndex);
                excuteCallBack(baseCallBack, excelRow);
            }
            curRowNum = 0;
            valueList = new ArrayList<ExcelCell>();
        }

    };

    /**
     * @param file 文件
     * @throws Exception IOException
     */
    public ExcelXlsxStream(File file) throws Exception {
        super.fileStream = new FileInputStream(file);
    }

    /**
     * @param fileStream 文件流
     */
    public ExcelXlsxStream(InputStream fileStream) {
        this.fileStream = fileStream;
    }

    /**
     * 解析cell编号
     *
     * @param cellNo 编号
     * @return String[]
     */
    private static String[] parseCellNo(String cellNo) {
        String[] cellNoArr = new String[2];
        for (int i = 0; i < cellNo.length(); i++) {
            char ch = cellNo.charAt(i);
            if (Character.isDigit(ch)) {
                cellNoArr[0] = cellNo.substring(0, i);
                cellNoArr[1] = cellNo.substring(i);
                break;
            }
        }
        return cellNoArr;
    }

    /**
     * 列字母转列数
     *
     * @param colStr 列字母
     * @return short
     */
    private static short excelColStrToNum(String colStr) {
        int len = colStr.length();
        short num = 0;
        short result = 0;
        for (int i = 0; i < len; i++) {
            char ch = colStr.charAt(len - i - 1);
            num = (short) (ch - 'A' + 1);
            num *= Math.pow(26, i);
            result += num;
        }
        return result;
    }

    /**
     * @throws Exception IOException
     */
    public void close() throws Exception {

        try {
            if (is != null) {
                is.close();
                is = null;
            }
        } catch (Exception e) {
            throw e;
        } finally {
            try {
                if (sst != null) {//新版本启用此方法
                    sst = null;
                }
            } catch (Exception e) {
                throw e;
            } finally {
                try {
                    if (pkg != null) {
                        pkg.close();
                        pkg = null;
                    }
                } catch (Exception e) {
                    throw e;
                } finally {
                    try {
                        if (fileStream != null) {
                            fileStream.close();
                            fileStream = null;
                        }
                    } catch (IOException e) {
                        throw e;
                    } finally {
                        defaultHandler = null;
                        r = null;
                        parser = null;
                        sheets = null;
                    }
                }

            }
        }
    }

    /**
     * 行事件处理方法
     *
     * @param baseCallBack 回调函数
     */
    public void rowStream(BaseCallBack baseCallBack) throws Exception {
        this.baseCallBack = baseCallBack;

        parser = SAXParserFactory.newInstance().newSAXParser().getXMLReader();
        parser.setContentHandler(defaultHandler);
        pkg = OPCPackage.open(fileStream);
        r = new XSSFReader(pkg);
        sst = r.getSharedStringsTable();
        sheets = (SheetIterator) r.getSheetsData();


        if (sheetIndexArr == null || sheetIndexArr[0] == null) {
            while (sheets.hasNext()) {
                is = sheets.next();
                sheetName = sheets.getSheetName();
                sheetList.add(sheetName);
                parser.parse(new InputSource(is));
                sheetIndex++;
            }
        } else {
            while (sheets.hasNext()) {
                is = sheets.next();
                sheetName = sheets.getSheetName();
                for (int index : sheetIndexArr) {
                    if (index == sheetIndex) {
                        parser.parse(new InputSource(is));
                    }
                }
                sheetList.add(sheetName);
                sheetIndex++;
            }
        }

    }

    /**
     * 指定工作簿
     *
     * @param sheetIndexArr 索引数组
     * @return BaseExcelStream
     */
    public ExcelEventStream sheetAt(Integer... sheetIndexArr) {
        this.sheetIndexArr = sheetIndexArr;
        return this;
    }

    public String getSheetName() {
        return sheetName;
    }

    public short getSheetIndex() {
        return sheetIndex;
    }
}