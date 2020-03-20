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

package ed.biodare.data.excel;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;
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


class XLSX2TextConverter {
    
    public void convert(Path inFile, Path outFile) throws InvalidFormatException, IOException {
        convert(inFile, outFile, ",");
    }

    public void convert(Path inFile, Path outFile, String SEP) throws InvalidFormatException, IOException {
        
        try (OPCPackage opcPackage = OPCPackage.open(inFile.toFile(), PackageAccess.READ)) {
            
            XSSFReader xssfReader = new XSSFReader(opcPackage);
            Iterator<InputStream> sheets = xssfReader.getSheetsData();
            if (!sheets.hasNext()) {
                throw new InvalidFormatException("Could not read any sheets");
            }
            
            try (BufferedWriter out = Files.newBufferedWriter(outFile)) {
                XMLReader sheetParser = SAXHelper.newXMLReader(); 
                XSSFSheetXMLHandler handler = buildXMLHandler(opcPackage, xssfReader, out, SEP);
                sheetParser.setContentHandler(handler);

                InputSource sheetSource = new InputSource(sheets.next());            
                sheetParser.parse(sheetSource);            
            }
        } catch (OpenXML4JException| SAXException | ParserConfigurationException e) {
            throw new IOException("Could not open the file: "+e.getMessage(),e);
        }
    }
    
    
    XSSFSheetXMLHandler buildXMLHandler(OPCPackage opcPackage, XSSFReader xssfReader, BufferedWriter out,
            String SEP) throws IOException, InvalidFormatException {
          
        try {
            StylesTable styles = new StylesTable(); //xssfReader.getStylesTable();
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(opcPackage);
            DataFormatter formatter = new NoFormatingDataFormatter();


            SheetContentsHandler sheetHandler = new SheetToCSVHandler(out, SEP);

            XSSFSheetXMLHandler handler = new XSSFSheetXMLHandler(
                      styles, null, strings, sheetHandler, formatter, false);            

            return handler;
        } catch (SAXException e) {
            throw new IOException("Could not parse: "+e.getMessage(),e);
        }
    }
    
    static class SheetToCSVHandler implements SheetContentsHandler {

        final BufferedWriter out;
        final String SEP;
        
        int currentRow = -1;
        int nextCol = 0; 
        List<Object> row = new ArrayList<>();
        
        SheetToCSVHandler(BufferedWriter out) {
            this(out, ",");
        }
        
        SheetToCSVHandler(BufferedWriter out, String SEP) {
            this.out = out;
            this.SEP = SEP;
        }
        
        @Override
        public void startRow(int rowNum) {
            
            // in case of rows gaps
            addRows(rowNum-currentRow-1);
            
            currentRow = rowNum;
            nextCol = 0;            
            row = new ArrayList<>();
        }

        @Override
        public void endRow(int i) {
            try {
                String line = row.stream()
                            .map( v -> v == null ? "" : v.toString())
                            .collect(Collectors.joining(SEP));
                out.write(line);                
                out.newLine();
            } catch (IOException e) {
                throw new RuntimeIOException(e);
            }
        }

        @Override
        public void cell(String cellAddress, String value, XSSFComment comment) {
            
            // gracefully handle missing CellRef here in a similar way as XSSFCell does
            if(cellAddress == null) {
                cellAddress = new CellAddress(currentRow, nextCol).formatAsString();
            }
            
            CellReference cellRef = new CellReference(cellAddress);
            int cellCol = cellRef.getCol();
            
            for (; nextCol < cellCol; nextCol++) {
                row.add(null);
            }
            
            if (value == null) value = "";
            if (value.contains(SEP)) value = "\""+value+"\"";
            
            row.add(value);
            nextCol++;
        }

        @Override
        public void endSheet() {
            
            try (out) {
                out.flush();
            } catch (IOException e) {
                throw new RuntimeIOException(e);
            }
        }
        
        void addRows(int number) {
            try {
                for (int i=0; i<number; i++) {
                    out.newLine();
                }
            } catch (IOException e) {
                throw new RuntimeIOException(e);
            }
        }        
        
    }
    
    public static class RuntimeIOException extends RuntimeException {

        public RuntimeIOException(IOException cause) {
            super(cause.getMessage(),cause);
        }
        
    }
    

}
