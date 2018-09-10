/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package ed.synthsys.util.excel;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.temporal.Temporal;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.NotOLE2FileException;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Object that simplifies acces to tabular data in the excel files.
 * @author tzielins
 */
public class ModernExcelView implements AutoCloseable {
    
    
    public static final CellCaster<String> STRING_CASTER = new StringCellCaster();
    public static final CellCaster<Double> DOUBLE_CASTER = new DoubleCellCaster();
    public static final CellCaster<Date> DATE_CASTER = new DateCellCaster();
    public static final CellCaster<Temporal> TEMPORAL_CASTER = new TemporalCellCaster();
    
    /**
     * Workbook which this object represents
     */
    final Workbook workbook;
    /**
     * Current sheet on which all the operations are performed
     */
    Sheet sheet;
    /**
     * Evalulator which will be used to calculate values of the formulas
     */
    FormulaEvaluator formEval;

    
    /**
     * Creates new excel view which is based on the content of the file.
     * @param file excel file to be read
     * @throws IOException when problems with io operations
     * @throws ExcelFormatException fi the file is not an excel file.
     */
    public ModernExcelView(Path file) throws IOException, ExcelFormatException  {
        
        this(file.toFile());
    }    
    
    /**
     * Creates new excel view which is based on the content of the file.
     * @param file excel file to be read
     * @throws IOException when problems with io operations
     * @throws ExcelFormatException fi the file is not an excel file.
     */
    public ModernExcelView(File file) throws IOException, ExcelFormatException  {
        
        try {
            this.workbook = WorkbookFactory.create(file,null,true);
            this.formEval = this.workbook.getCreationHelper().createFormulaEvaluator();
            this.formEval.setIgnoreMissingWorkbooks(true);
            
            selectSheet(0);
        } catch (InvalidFormatException | IllegalArgumentException | NotOLE2FileException e) {
            throw new ExcelFormatException("Not valid excel fle: "+e.getMessage(),e);
        }
    }
    
    /**
     * Creates new excel view which is based on the content of the file.
     * @param in input stream with excel content to be read
     * @throws IOException when problems with io operations
     * @throws ExcelFormatException fi the file is not an excel file.
     */
    public ModernExcelView(InputStream in) throws IOException, ExcelFormatException  {
        
        try {
            this.workbook = WorkbookFactory.create(in);
            this.formEval = this.workbook.getCreationHelper().createFormulaEvaluator();
            this.formEval.setIgnoreMissingWorkbooks(true);
            
            selectSheet(0);
        } catch (InvalidFormatException | IllegalArgumentException | NotOLE2FileException e) {
            throw new ExcelFormatException("Not valid excel: "+e.getMessage(),e);
        }
    }
    
    
    
    @Override
    public void close() {
        try {
            if (this.formEval != null) {
                this.formEval.clearAllCachedResultValues();
            }
            this.formEval = null;
            this.workbook.close();
        } catch (IOException e) {
            throw new WorkbookCloseException("Could not close workbook: "+e.getMessage(),e);
        }
    }
    
    /**
     * Checks if the file represents a valid excel file.
     * @param file to be checked
     * @return true if the file is readable excel
     * @throws IOException if IO problems.
     */
    public static boolean isExcelFile(Path file) throws IOException {
        return isExcelFile(file.toFile());
    }    
    
    /**
     * Checks if the file represents a valid excel file.
     * @param file to be checked
     * @return true if the file is readable excel
     * @throws IOException if IO problems.
     */
    public static boolean isExcelFile(File file) throws IOException {

        try {

            Workbook wr = WorkbookFactory.create(file);
            if (wr == null) return false;
            if (wr.getNumberOfSheets() < 1) return false;
            Sheet sh = wr.getSheetAt(0);
            return sh != null;
        } catch (    InvalidFormatException | IllegalArgumentException | NotOLE2FileException e)  {
            return false;
        }
    }

    /**
     * Changes the active sheet to the given one.
     * @param i sheet number, 0-based
     */
    public void selectSheet(int i) {
        sheet = workbook.getSheetAt(i);
    }

    /**
     * Reads content of the row till the end of the row
     * @param rowNr 0-based row number
     * @param firstCol 0-based number of the column from which to start
     * @return list of read values, with nulls in case of empty cell
     */
    public List<String> readStringRow(int rowNr,int firstCol) {
        return readRow(rowNr,firstCol,new StringCellCaster());
    }
    
    /**
     * Reads content of the row in the given range of columns
     * @param rowNr 0-based row number
     * @param firstCol 0-based number of the column from which to start
     * @param lastCol 0-based last column to read from, inclusive
     * @return list of read values, with nulls in case of empty cell
     */
    public List<String> readStringRow(int rowNr,int firstCol,int lastCol) {
        Row row = getRow(rowNr);
        if (row == null) return Collections.emptyList();
        return readRow(row,firstCol,lastCol,new StringCellCaster());
    }
    
    /**
     * 
     * @param rowNr
     * @return row of the given nr or null if not found
     */
    protected final Row getRow(int rowNr) {
        Row row = sheet.getRow(rowNr);
        //if (row == null) throw new RobustProcessException("No row nr: "+rowNr);
        return row;
    }
        
    /**
     * Reads the row content and cast the values to the requested type as handled by the caster
     * @param <T> type of the elements to which the read values should be cast
     * @param rowNr 0-based row number
     * @param firstCol 0-based number of the column from which to start
     * @param caster code that can convert excel read value to the requested type
     * @return list of read values with null for missing/or not convertable values.
     */
    public <T> List<T> readRow(int rowNr, int firstCol,CellCaster<T> caster) {
        
        Row row = getRow(rowNr);
        if (row == null) return Collections.emptyList();
        int lastCol = row.getLastCellNum()-1;
        return readRow(row,firstCol,lastCol,caster);
    }
    

    /**
     * Reads the row content and cast the values to the requested type as handled by the caster
     * @param <T> type of the elements to which the read values should be cast
     * @param row row to read from
     * @param firstCol 0-based number of the column from which to start
     * @param lastCol 0-based last column to read from, inclusive
     * @param caster code that can convert excel read value to the requested type
     * @return list of read values with null for missing/or not convertable values.
     */
    public <T> List<T> readRow(Row row, int firstCol, int lastCol,CellCaster<T> caster)  {
        if (lastCol < firstCol) throw new IllegalArgumentException("Wrong column: "+firstCol+"-"+lastCol);
        List<T> list = new ArrayList<>(lastCol-firstCol);
        
        for (int col = firstCol;col<=lastCol;col++) {
            Cell cell = row.getCell(col,Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            list.add(caster.cast(cell, formEval));
        }
        return list;
    }

    /**
     * Number of the last column that exists in the given row.
     * @param rowNr 0-based row number to check
     * @return
     */
    public int getLastColumn(int rowNr) {
        Row row = getRow(rowNr);
        if (row == null) throw new IllegalArgumentException("No row nr: "+rowNr);
        return row.getLastCellNum()-1;
    }
    
    /**
     * Gets last row number on the sheet
     * @return rowNr 0-based
     */
    public int getLastRow() {
        return sheet.getLastRowNum();
    }

    /**
     * Reads double values from one column in the active sheet.
     * @param colNr 0-based column number
     * @param firstRow 0-based number of row from which to read from
     * @return list of double that corresponds to all the rows till the end, with null for missing or numercial values
     */
    public List<Double> readDoubleColumn(int colNr, int firstRow) {
        
        return readColumn(colNr,firstRow,new DoubleCellCaster());
    }
    
    /**
     * Reads double values from one column in the active sheet.
     * @param colNr 0-based column number
     * @param firstRow 0-based number of row from which to read from
     * @param lastRow 0-based last row with data (inclusive)
     * @return list of double that corresponds to all the rows from fist till last, with null for missing or numercial values
     */
    public List<Double> readDoubleColumn(int colNr, int firstRow,int lastRow)  {
        
        return readColumn(colNr,firstRow,lastRow,new DoubleCellCaster());
    }
    
    
    public <T> List<T> readColumn(int colNr,int firstRow,CellCaster<T> caster)  {
        
        int lastRow = sheet.getLastRowNum();
        return readColumn(colNr,firstRow,lastRow,caster);
    }

    public <T> List<T> readColumn(int colNr,int firstRow,int lastRow,CellCaster<T> caster) throws IllegalArgumentException {
        
        if (lastRow< firstRow) throw new IllegalArgumentException("Wrong rows range: "+firstRow+"-"+lastRow);
        List<T> list = new ArrayList<>(lastRow-firstRow+1);
        
        for (int rowIx = firstRow;rowIx<=lastRow;rowIx++) {
            Row row = sheet.getRow(rowIx);
            if (row == null) {
                list.add(caster.cast(null, formEval));
                continue;
            }
            Cell cell = row.getCell(colNr, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            list.add(caster.cast(cell, formEval));
        }
        return list;
    }
    
    /**
     * Reads the columns in the given range of columns and rows
     * @param <T> type of the values to be read, as determined by the caster
     * @param firstCol 0-based first column to read from
     * @param lastCol 0-based last column to read from (inclusive)
     * @param firstRow 0-based frist row from which data will be read
     * @param lastRow 0-base last row from which data will be read (inclusive)
     * @param caster converter of cell values to the required values
     * @return list of list, in which each list correspond to one data column, missing or not convertible values are represented as nulls
     */
    public <T> List<List<T>> readColumns(int firstCol,int lastCol,int firstRow,int lastRow,CellCaster<T> caster) {
        if (lastCol < firstCol) throw new IllegalArgumentException("Wrong columns range: "+firstCol+"-"+lastCol);
        
        List<List<T>> columns = new ArrayList<>(lastCol-firstCol+1);
        for (int colIx = firstCol;colIx<=lastCol;colIx++) {
            columns.add(readColumn(colIx, firstRow, lastRow,caster));
        }
        return columns;
    }

    public String readStringCell(int rowNr, int colNr) {
        return readCell(rowNr,colNr,STRING_CASTER);
    }
    
    public Date readDateCell(int rowNr, int colNr) {
        return readCell(rowNr,colNr,DATE_CASTER);
    }
    
    public Temporal readTemporalCell(int rowNr,int colNr) {
        return readCell(rowNr,colNr,TEMPORAL_CASTER);
    }
    
    public Double readDoubleCell(int rowNr,int colNr) {
        return readCell(rowNr,colNr,DOUBLE_CASTER);
    }
    
    

    public <T> T readCell(int rowNr, int colNr, CellCaster<T> caster) {
        Row row = sheet.getRow(rowNr);
        if (row == null) return caster.cast(null, formEval);
        Cell cell =row.getCell(colNr, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        return caster.cast(cell, formEval);
    }

    /**
     * Finds parameter value in the given range of search rows. Parameters are represented
     * as pair of cells, in which first cell (column 0) is the parameter name, and the second cell (column 1)
     * is the parameter value. Mothods scans the row in the requested range starting with the firstOne and finishing with the lastRow,
     * returning first value for which cell content at column 0 matches the paramName
     * @param <T> type for the return value which is set by the CellCaster type 
     * @param paramName name of the parameter (value of the first cell in a row)
     * @param firstRow 0-based first row from which data will be read
     * @param lastRow 0-base last row from which data will be read (inclusive)
     * @param caster converter of cell values to the required values
     * @return found value for the given parameters or null if such parameter name could not be found
     */
    public <T> T findParam(String paramName,int firstRow,int lastRow,CellCaster<T> caster) {

        for (int i = firstRow;i<=lastRow;i++) {
            String pN = readCell(i, 0, STRING_CASTER);
            if (paramName.equals(pN))
                return readCell(i,1,caster);
        }
        return caster.cast(null, formEval);
    }
    
    public String findParamAsString(String paramName,int firstRow,int lastRow) {

        return findParam(paramName, firstRow, lastRow, STRING_CASTER);
    }
    
    public String getCurrentSheetName() {
        return sheet.getSheetName();
    }
    
    public int getCurrentSheetNr() {
        return workbook.getSheetIndex(sheet);
    }
    
    
    public static interface CellCaster<T> {
        public T cast(Cell cell,FormulaEvaluator formEval);
    }
    
    protected final static boolean isMathInteger(double val) {
        return Math.rint(val) == val;
    }
    
   
    protected static class StringCellCaster implements CellCaster<String> {

        @Override
        public String cast(Cell cell, FormulaEvaluator formEval) {
            if (cell == null) return null;            
            switch(cell.getCellType()) {
                case Cell.CELL_TYPE_STRING: return cell.getRichStringCellValue().getString().trim();
                case Cell.CELL_TYPE_NUMERIC: {
                    final double val = cell.getNumericCellValue();
                    if (isMathInteger(val)) return Long.toString(Math.round(val));
                    return ""+cell.getNumericCellValue();
                }
                case Cell.CELL_TYPE_BOOLEAN: return ""+cell.getBooleanCellValue();
                case Cell.CELL_TYPE_FORMULA: {
                    //logger.debug("Formula in #0,#1 :#2",""+cell.getRowIndex(),""+cell.getColumnIndex(),cell.getCellFormula());
                    try {
                        CellValue val = formEval.evaluate(cell);
                        if (val.getCellType() == Cell.CELL_TYPE_NUMERIC) return ""+val.getNumberValue();
                        if (val.getCellType() == Cell.CELL_TYPE_STRING) return val.getStringValue().trim();
                        if (val.getCellType() == Cell.CELL_TYPE_BOOLEAN) return ""+val.getBooleanValue();
                    } catch (FormulaParseException e) {
                        return null;
                    }
                    return null;
                }
                default: return null;
            }            
        }        
    }
    
    protected static class DoubleCellCaster implements  CellCaster<Double> {

        @Override
        public Double cast(Cell cell, FormulaEvaluator formEval) {
            if (cell == null) return null;
            
            switch(cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC: return cell.getNumericCellValue();
                case Cell.CELL_TYPE_FORMULA: {
                    //logger.debug("Formula in #0,#1 :#2",""+cell.getRowIndex(),""+cell.getColumnIndex(),cell.getCellFormula());
                    try {
                        CellValue val = formEval.evaluate(cell);
                        if (val.getCellType() == Cell.CELL_TYPE_NUMERIC) return val.getNumberValue();
                    } catch (FormulaParseException e) {
                        //logger.warn("Error evaluating formula: #0, #1", e, cell.getCellFormula(),e.getMessage());
                        return null;
                    }
                    return null;
                }
                case Cell.CELL_TYPE_STRING: {
                    try {
                        return Double.parseDouble(cell.getStringCellValue());
                    } catch(Exception e) {
                        return null;
                    }
                }
                default: return null;
            }            
        }
        
    }

    protected static class DateCellCaster implements  CellCaster<Date> {

        @Override
        public Date cast(Cell cell, FormulaEvaluator formEval) {
            if (cell == null) return null;
            
            switch(cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC: return cell.getDateCellValue();
                case Cell.CELL_TYPE_FORMULA: {
                    //logger.debug("Formula in #0,#1 :#2",""+cell.getRowIndex(),""+cell.getColumnIndex(),cell.getCellFormula());
                    try {
                        CellValue val = formEval.evaluate(cell);
                        
                        //if (val.getCellType() == Cell.CELL_TYPE_NUMERIC) return val.getDateCellValue();
                        
                    } catch (FormulaParseException e) {
                        //logger.warn("Error evaluating formula: #0, #1", e, cell.getCellFormula(),e.getMessage());
                        return null;
                    }
                    return null;
                }
                case Cell.CELL_TYPE_STRING: {
                    try {
                        //return Double.parseDouble(cell.getStringCellValue());
                        return null;
                    } catch(Exception e) {
                        return null;
                    }
                }
                default: return null;
            }            
        }
        
    }
    
    protected static class TemporalCellCaster implements  CellCaster<Temporal> {

        @Override
        public Temporal cast(Cell cell, FormulaEvaluator formEval) {
            Date date = DATE_CASTER.cast(cell, formEval);
            if (date == null) return null;
            return LocalDateTime.ofInstant(date.toInstant(),ZoneId.systemDefault());
            
        }
    }
    
    public static class WorkbookCloseException extends RuntimeException {
        WorkbookCloseException(String msg,Throwable e) {
            super(msg,e);
        }
    }
}
