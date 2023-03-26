/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package com.dc.eventpoi.test.temp;

import java.util.List;

import com.alibaba.fastjson.JSON;
import com.dc.eventpoi.test.temp.read.CellReadCallBack;
import com.dc.eventpoi.test.temp.read.XlsxReadStream;


/**
 * XSSF and SAX (Event API) basic example.
 * See {@link XLSX2CSV} for a fuller example of doing
 *  XSLX processing with the XSSF Event code.
 */
public class TestReadXlsx {
	public static void main(String[] args) throws Throwable  {
    	XlsxReadStream howto = new XlsxReadStream();
    	howto.setReadSheetIndex(0);
    	howto.setFileName("E:\\eclipse-workspace-2022-12\\excel-eventpoi\\my_test_temp\\file.xlsx");
//    	howto.registerCallBack(new ReadCallBack() {
//			
//			@Override
//			public void callBack(StreamReadBaseCallBack baseCallBack) {
//				System.err.println(JSON.toJSONString(baseCallBack));				
//			}
//		});
    	
    	howto.registerCallBack(new RowCallBack() {
			@Override
			public void callBack(int rowIndex, List<CellReadCallBack> cellList) {
				System.err.println("rowIndex="+rowIndex+",list="+JSON.toJSONString(cellList));
			}
		});
    	howto.doRead();
        //howto.processFirstSheet(args[0]);
        
    }
	
}
