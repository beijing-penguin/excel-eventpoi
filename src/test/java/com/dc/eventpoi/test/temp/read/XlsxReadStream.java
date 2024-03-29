package com.dc.eventpoi.test.temp.read;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.util.StringUtil;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.model.SharedStrings;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import com.dc.eventpoi.test.temp.CellCallBack;
import com.dc.eventpoi.test.temp.RegCallBack;
import com.dc.eventpoi.test.temp.RowCallBack;

public class XlsxReadStream {

	private RegCallBack regCallBack;
	private InputStream fileInputStream;
	private String fileName;
	private Integer readSheetIndex = 0;
	private Integer sheetIndex = 0;

	private List<CellReadCallBack> tempCellReadCallBackList = new ArrayList<>();

	private List<CellReadCallBack> rtCellList = new ArrayList<>();
	
	private LinkedHashMap<SheetReadCallBack,List<CellReadCallBack>> dataMap = new LinkedHashMap<SheetReadCallBack,List<CellReadCallBack>>();
	
	public void callBack(StreamReadBaseCallBack baseCallBack) {
		if(baseCallBack instanceof CellReadCallBack) {
			CellReadCallBack cellReadCallBack = (CellReadCallBack)baseCallBack;

			String[] cellNoArr = this.parseCellNo(cellReadCallBack.getCellNo());
			short cellNum = (short) (excelColStrToNum(cellNoArr[0]) - 1);
			int rowNum = Integer.parseInt(cellNoArr[1]) - 1;

			cellReadCallBack.setCellIndex(cellNum);
			cellReadCallBack.setRowIndex(rowNum);

			if(regCallBack != null) {
				if(regCallBack instanceof RowCallBack) {
					RowCallBack rowCallBack = (RowCallBack)regCallBack;
					if(tempCellReadCallBackList.size() == 0) {
						tempCellReadCallBackList.add(cellReadCallBack);
					}else {
						if(tempCellReadCallBackList.get(0).getRowIndex() != cellReadCallBack.getRowIndex()) {
							rowCallBack.callBack(tempCellReadCallBackList.get(0).getRowIndex(), tempCellReadCallBackList);
							tempCellReadCallBackList.clear();
							tempCellReadCallBackList.add(cellReadCallBack);
						}else {
							tempCellReadCallBackList.add(cellReadCallBack);
						}
					}
				}else if(regCallBack instanceof CellCallBack) {
					CellCallBack rowCallBack = (CellCallBack)regCallBack;
					rowCallBack.callBack(cellReadCallBack);
				}
			}else {
				rtCellList.add(cellReadCallBack);
			}
		}

	}

	/**
	 * 列字母转列数
	 *
	 * @param colStr 列字母
	 * @return short
	 */
	private short excelColStrToNum(String colStr) {
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

	public String[] parseCellNo(String cellNo) {
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

	public void processAllSheets() throws Throwable {
		OPCPackage pkg = null;
		if(fileName != null) {
			pkg = OPCPackage.open(fileName, PackageAccess.READ);
		}else if(fileInputStream != null) {
			pkg = OPCPackage.open(fileInputStream);
		}
		
		try {
			XSSFReader r = new XSSFReader(pkg);
			SharedStrings sst = r.getSharedStringsTable();

			XMLReader parser = fetchSheetParser(sst);

			SheetIterator sheets = (SheetIterator)r.getSheetsData();
			while (sheets.hasNext()) {
				try (InputStream sheet = sheets.next()) {
					if(readSheetIndex != null && sheetIndex != readSheetIndex) {
						sheetIndex++;
						continue;
					}
					SheetReadCallBack sheetCallBack = new SheetReadCallBack();
					sheetCallBack.setSheetIndex(sheetIndex);
					sheetCallBack.setSheetName(sheets.getSheetName());
					this.callBack(sheetCallBack);

					//解析excel
					InputSource sheetSource = new InputSource(sheet);
					parser.parse(sheetSource);
					sheetIndex++;

					if(regCallBack != null) {
						if(regCallBack instanceof RowCallBack) {
							if(tempCellReadCallBackList.size() > 0 ) {
								RowCallBack rowCallBack = (RowCallBack)regCallBack;
								rowCallBack.callBack(tempCellReadCallBackList.get(0).getRowIndex(), tempCellReadCallBackList);
								tempCellReadCallBackList.clear();
							}
						}
					}else {
						dataMap.put(sheetCallBack, rtCellList);
						rtCellList = new ArrayList<>();
					}
				}
			}
			
			
		}catch (Throwable e) {
			throw e;
		}finally {
			try {
				if(pkg != null) {
					pkg.close();
				}
			}catch (Throwable e) {
				throw e;
			}finally {
				if(fileInputStream != null) {
					fileInputStream.close();
				}
			}
			sheetIndex = 0;
			readSheetIndex = 0;
			regCallBack = null;
		}
	}

	public XMLReader fetchSheetParser(SharedStrings sst) throws SAXException, ParserConfigurationException {
		XMLReader parser = XMLHelper.newXMLReader();
		ContentHandler handler = new SheetHandler(sst,this);
		parser.setContentHandler(handler);
		return parser;
	}

	/**
	 * See org.xml.sax.helpers.DefaultHandler javadocs
	 */
	private static class SheetHandler extends DefaultHandler {
		private XlsxReadStream xlsxReadStream;
		private String tempCellNo;
		private final SharedStrings sst;
		private String lastContents;
		private boolean nextIsString;
		private boolean inlineStr;
		private final LruCache<Integer,String> lruCache = new LruCache<>(50);

		private static class LruCache<A,B> extends LinkedHashMap<A, B> {
			private static final long serialVersionUID = 1L;
			private final int maxEntries;

			public LruCache(final int maxEntries) {
				super(maxEntries + 1, 1.0f, true);
				this.maxEntries = maxEntries;
			}

			@Override
			protected boolean removeEldestEntry(final Map.Entry<A, B> eldest) {
				return super.size() > maxEntries;
			}
		}

		private SheetHandler(SharedStrings sst,XlsxReadStream xlsxReadStream) {
			this.sst = sst;
			this.xlsxReadStream = xlsxReadStream;
		}

		@Override
		public void startElement(String uri, String localName, String name,
				Attributes attributes) throws SAXException {
			// c => cell
			if(name.equals("c")) {
				// Print the cell reference
				tempCellNo = attributes.getValue("r");
				// Figure out if the value is an index in the SST
				String cellType = attributes.getValue("t");
				nextIsString = cellType != null && cellType.equals("s");
				inlineStr = cellType != null && cellType.equals("inlineStr");
			}
			// Clear contents cache
			lastContents = "";
		}

		@Override
		public void endElement(String uri, String localName, String name)
				throws SAXException {
			// Process the last contents as required.
			// Do now, as characters() may be called more than once
			if(nextIsString && StringUtil.isNotBlank(lastContents)) {
				Integer idx = Integer.valueOf(lastContents);
				lastContents = lruCache.get(idx);
				if (lastContents == null && !lruCache.containsKey(idx)) {
					lastContents = sst.getItemAt(idx).getString();
					lruCache.put(idx, lastContents);
				}
				nextIsString = false;
			}

			// v => contents of a cell
			// Output after we've seen the string contents
			if(name.equals("v") || (inlineStr && name.equals("c"))) {
				CellReadCallBack callBack = new CellReadCallBack();
				callBack.setCellNo(tempCellNo);
				callBack.setCellValue(lastContents);
				xlsxReadStream.callBack(callBack);
			}
		}

		@Override
		public void characters(char[] ch, int start, int length) throws SAXException { // NOSONAR
			lastContents += new String(ch, start, length);
		}
	}


	public Integer getReadSheetIndex() {
		return readSheetIndex;
	}

	public void setReadSheetIndex(Integer readSheetIndex) {
		this.readSheetIndex = readSheetIndex;
	}


	public String getFileName() {
		return fileName;
	}

	public XlsxReadStream setFileName(String fileName) {
		this.fileName = fileName;
		return this;
	}
	public InputStream getFileInputStream() {
		return fileInputStream;
	}

	public XlsxReadStream setFileInputStream(InputStream fileInputStream) {
		this.fileInputStream = fileInputStream;
		return this;
	}

	
	public LinkedHashMap<SheetReadCallBack, List<CellReadCallBack>> getDataMap() {
		return dataMap;
	}

	public void setDataMap(LinkedHashMap<SheetReadCallBack, List<CellReadCallBack>> dataMap) {
		this.dataMap = dataMap;
	}

	public void doRead(RegCallBack regCallBack) throws Throwable {
		this.regCallBack = regCallBack;
		this.processAllSheets();
	}
	public LinkedHashMap<SheetReadCallBack,List<CellReadCallBack>> doRead() throws Throwable {
		this.doRead(null);
		return dataMap;
	}
}
