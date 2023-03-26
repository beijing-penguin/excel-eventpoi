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

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dc.eventpoi.test.Me;

/**
 * Demonstrates how to insert pictures in a SpreadsheetML document
 */
public final class WorkingWithPictures {
    private WorkingWithPictures() {}

    public static void main(String[] args) throws IOException {

        //create a new workbook
        try (Workbook wb = new XSSFWorkbook()) {
            CreationHelper helper = wb.getCreationHelper();

            //add a picture in this workbook.
            String img_file_path = new File(Me.class.getResource("unnamed.jpg").getPath()).getAbsolutePath();
    		byte[] bytes = Files.readAllBytes(Paths.get(img_file_path));
            int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);

            //create sheet
            Sheet sheet = wb.createSheet();

            //create drawing
            Drawing<?> drawing = sheet.createDrawingPatriarch();

            //add a picture shape
            ClientAnchor anchor = helper.createClientAnchor();
            anchor.setCol1(1);
            anchor.setRow1(1);
            Picture pict = drawing.createPicture(anchor, pictureIdx);

            //auto-size picture
            pict.resize(0.1);

            //save workbook
            String file = "my_test_temp/picture.xlsx";
            try (OutputStream fileOut = new FileOutputStream(file)) {
                wb.write(fileOut);
            }
        }
    }
}
