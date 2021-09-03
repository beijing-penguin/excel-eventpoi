package com.dc.eventpoi.core;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BlankRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.EOFRecord;
import org.apache.poi.hssf.record.LabelRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.RowRecord;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.record.SelectionRecord;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * 事件流xlx处理
 * @author beijing-penguin
 * @date: 2019年1月16日
 * @eg:https://svn.apache.org/repos/asf/poi/trunk/poi-examples/src/main/java/org/apache/poi/examples/hssf/usermodel/EventExample.java
 * @mainUrl:http://poi.apache.org/components/spreadsheet/examples.html
 */
public class ExcelXlsStream extends BaseExcelStream implements ExcelEventStream{
	
	private static Log LOG = LogFactory.getLog(ExcelXlsStream.class);
	
	/**
	 * 
	 */
	private POIFSFileSystem poifs;
	/**
	 * 
	 */
	private InputStream din;
	/**
	 * 
	 */
	private HSSFRequest req;
	/**
	 * 
	 */
	private HSSFEventFactory factory;
	/**
	 * 
	 */
	private BaseCallBack baseCallBack;
	/**
	 * 
	 */
	private short sheetIndex = 0;
	/**
	 * 
	 */
	private HSSFListener listener =  new HSSFListener() {
		private byte eofNum = 0;
		private int curRowIndex = 0;
		private List<ExcelCell> valueList = new ArrayList<ExcelCell>();

		/**
		 * 
		 */
		private SSTRecord sstrec;

		public void eventHandle(int rowIndex,ExcelCell excelCell) {
			/*if(rowIndex==8) {
                System.err.println(11);
            }*/
			if(excelCell==null) {
				if(allowSheet(sheetIndexArr, sheetIndex)) {
					ExcelRow excelRow= new ExcelRow();
					excelRow.setRowIndex(excelRow.getRowIndex());
					excuteCallBack(baseCallBack,excelRow);
				}
				sheetIndex++;
				curRowIndex = 0;
				valueList = new ArrayList<ExcelCell>();
			}

			if(rowIndex==curRowIndex) {
				valueList.add(excelCell);
			}else {
				ExcelRow excelRow= new ExcelRow();
				if(allowSheet(sheetIndexArr, sheetIndex)) {
					excelRow.setCellList(valueList);
					excelRow.setRowIndex(curRowIndex);
					excuteCallBack(baseCallBack,excelRow);
					//rowCallBack.getRow(excelRow);
				}
				valueList = new ArrayList<ExcelCell>();
				valueList.add(excelCell);
				curRowIndex = rowIndex;
			}
		}
		@Override
		public void processRecord(Record record) {
			//System.out.println(record.getSid());
			switch (record.getSid()) {
			// the BOFRecord can represent either the beginning of a sheet or the workbook
			case BOFRecord.sid:
				/*BOFRecord bof = (BOFRecord) record;
                if (bof.getType() == BOFRecord.TYPE_WORKBOOK) {
                } else if (bof.getType() == BOFRecord.TYPE_WORKSHEET) {
                }*/
				break;
			case BoundSheetRecord.sid:
				BoundSheetRecord bsr = (BoundSheetRecord) record;
				sheetList.add(bsr.getSheetname());
				break;
			case RowRecord.sid:
				//RowRecord rowrec = (RowRecord) record;
				//System.out.println("Row found, first column at " + rowrec.getFirstCol() + " last column at " + rowrec.getLastCol());
				break;
			case NumberRecord.sid:
				NumberRecord numrec = (NumberRecord) record;
				int rowNum1 = numrec.getRow();
				short colNum1 = numrec.getColumn();
				DecimalFormat decimalFormat = new DecimalFormat("###################.###########");  
				this.eventHandle(rowNum1, new ExcelCell(colNum1, decimalFormat.format(numrec.getValue())));
				break;
				// SSTRecords store a array of unique strings used in Excel.
			case SSTRecord.sid:
				sstrec = (SSTRecord) record;
				/*for (int k = 0; k < sstrec.getNumUniqueStrings(); k++) {
                    System.out.println("String table value " + k + " = " + sstrec.getString(k));
                }*/
				break;
			case LabelSSTRecord.sid:
				LabelSSTRecord lrec = (LabelSSTRecord) record;
				int rowNum2 = lrec.getRow();
				short colNum2 = lrec.getColumn();
				this.eventHandle(rowNum2,new ExcelCell(colNum2, sstrec.getString(lrec.getSSTIndex()).getString()));
				//System.err.println("第"+lrec.getRow()+"行，第"+lrec.getColumn()+"列，值="+sstrec.getString(lrec.getSSTIndex()));
				break;
			case LabelRecord.sid:  
				LabelRecord lr = (LabelRecord) record;
				int rowNum3 = lr.getRow();
				short colNum3 = lr.getColumn();
				this.eventHandle(rowNum3,new ExcelCell(colNum3, lr.getValue()));
				//System.err.println("sheet="+(sheetIndex--)+"第"+lr.getRow()+"行，第"+lr.getColumn()+"列，值="+lr.getValue());
				break; 
			case BlankRecord.sid:  
				break; 
			case SelectionRecord.sid:  
				//SelectionRecord sr = (SelectionRecord) record;
				//System.out.println(sr.getActiveCellRow());
				break; 
			case EOFRecord.sid:  
				//EOFRecord er = (EOFRecord) record;
				if((sheetIndex+1) == (eofNum)) {
					this.eventHandle(0,null);
				}
				eofNum++;
				break; 
			}
		}
	};
	/**
	 * 允许的sheet事件发生判断
	 * @param sheetIndexArr sheetIndexArr
	 * @param sheetIndex sheetIndex
	 * @return boolean
	 * @date 2019-01-16 13:58:00
	 */
	public boolean allowSheet(Integer[] sheetIndexArr,short sheetIndex) {
		if(sheetIndexArr==null) {
			return true;
		}
		for(int index : sheetIndexArr) {
			if(index == sheetIndex) {
				return true;
			}
		}
		return false;
	}
	/**
	 * 
	 * @param file 文件
	 * @throws Exception 
	 */
	public ExcelXlsStream(File file) throws Exception {
		//super.file = file;
		super.fileStream = new FileInputStream(file);
	}
	/**
	 * 
	 * @param fileStream 文件流
	 */
	public ExcelXlsStream(InputStream fileStream) {
		super.fileStream = fileStream;
	}
	/**
	 * 字节数据
	 * @param bytes 字节
	 */
	public ExcelXlsStream(byte[] bytes) {
		super.fileStream = new ByteArrayInputStream(bytes);
	}

	/**
	 * 
	 * @throws Exception 
	 * @date 2019-01-15 13:39:02
	 */
	public void close() throws Exception {
		try {
			if(fileStream!=null) {
				fileStream.close();
				fileStream = null;
			}
		}catch (Exception e) {
			LOG.error("",e);
			throw e;
		}finally {
			try {
				if(poifs!=null) {
					poifs.close();
					poifs = null;
				}
			}catch (Exception e) {
				LOG.error("",e);
				throw e;
			}finally {
				if(din!=null) {
					try {
						din.close();
					} catch (Exception e) {
						LOG.error("",e);
						throw e;
					}
				}
				req = null;
				factory = null;
			}
		}
	}

	/**
	 * 获取事件发生时的工作簿名称
	 * @return String
	 * @date 2019-01-15 10:59:53
	 */
	public String getSheetName() {
		return sheetList.get(sheetIndex);
	}

	/**
	 * 行回调
	 * @param baseCallBack 回调函数
	 * @throws Exception 
	 * @date 2019-01-14 18:06:24
	 */
	public void rowStream(BaseCallBack baseCallBack) throws Exception {
		this.baseCallBack = baseCallBack;
		this.processEvents();
	}
	/**
	 * 
	 * @throws Exception 
	 * @date 2019-01-15 11:19:06
	 */
	private void processEvents() throws Exception {
		try {
			poifs = new POIFSFileSystem(fileStream);
			din = poifs.createDocumentInputStream("Workbook");
			req = new HSSFRequest();
			req.addListenerForAllRecords(listener);
			factory = new HSSFEventFactory();
			factory.processEvents(req, din);
		}catch (Exception e) {
			throw e;
		}
	}
	@Override
	public short getSheetIndex() {
		return sheetIndex;
	}
	/**
	 * 指定工作簿
	 * @param sheetIndexArr 索引数组
	 * @return BaseExcelStream
	 * @date 2019-01-21 11:01:45
	 */
	public ExcelEventStream sheetAt(Integer... sheetIndexArr) {
		this.sheetIndexArr = sheetIndexArr;
		return this;
	}
}