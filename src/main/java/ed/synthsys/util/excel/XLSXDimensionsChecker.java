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

// from http://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/xssf/eventusermodel/

package ed.synthsys.util.excel;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Iterator;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

// Based on http://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/xssf/eventusermodel/


class XLSXDimensionsChecker {
    

    public int[] rowsColsDimensions(Path inFile) throws InvalidFormatException, IOException {
        
        try (OPCPackage opcPackage = OPCPackage.open(inFile.toFile(), PackageAccess.READ)) {
            
            XSSFReader xssfReader = new XSSFReader(opcPackage);
            Iterator<InputStream> sheets = xssfReader.getSheetsData();
            if (!sheets.hasNext()) {
                throw new InvalidFormatException("Could not read any sheets");
            }

            SheetToSizeHandler sizer = new SheetToSizeHandler();            
            XSSFSheetXMLHandler handler = buildXMLHandler(opcPackage, xssfReader, sizer);

            XMLReader sheetParser = SAXHelper.newXMLReader(); 
            sheetParser.setContentHandler(handler);

            InputSource sheetSource = new InputSource(sheets.next());            
            sheetParser.parse(sheetSource);            
            
            return new int[]{sizer.rows, sizer.cols};
        } catch (OpenXML4JException| SAXException | ParserConfigurationException e) {
            throw new IOException("Could not open the file: "+e.getMessage(),e);
        }
    }
    
    
    XSSFSheetXMLHandler buildXMLHandler(OPCPackage opcPackage, XSSFReader xssfReader, SheetContentsHandler sheetHandler) throws IOException, InvalidFormatException {
          
        try {
            StylesTable styles = new StylesTable(); // xssfReader.getStylesTable();
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(opcPackage);
            DataFormatter formatter = new NoFormatingDataFormatter();

            XSSFSheetXMLHandler handler = new XSSFSheetXMLHandler(
                      styles, null, strings, sheetHandler, formatter, false);            

            return handler;
        } catch (SAXException e) {
            throw new IOException("Could not parse: "+e.getMessage(),e);
        }
    }
    
    static class SheetToSizeHandler implements SheetContentsHandler {
        
        public int rows = 0;
        public int cols = 0; 
        
        SheetToSizeHandler() {
        }
        
        
        @Override
        public void startRow(int rowNum) {
            
            rows = Math.max(rows, rowNum+1);
        }

        @Override
        public void endRow(int i) {
        }

        @Override
        public void cell(String cellAddress, String value, XSSFComment comment) {
            
            if(cellAddress == null) {
                return;
            }
            
            CellReference cellRef = new CellReference(cellAddress);
            int cellCol = cellRef.getCol();
            
            cols = Math.max(cols, cellCol+1);
            
        }

        @Override
        public void endSheet() {
        }
        
        
    }
    
    

}
