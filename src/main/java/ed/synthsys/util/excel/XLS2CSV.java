/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ed.synthsys.util.excel;

import java.io.BufferedWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author Tomasz Zielinski <tomasz.zielinski@ed.ac.uk>
 */
class XLS2CSV {
    
    String SEP = ",";
    
    public void convert(Path inFile, Path outFile) throws IOException {
        
        Workbook workbook = WorkbookFactory.create(inFile.toFile());
        if (workbook.getNumberOfSheets() < 1) {
            throw new IOException("No sheets in file");
        }
        
        // FormulaEvaluator formEval = workbook.getCreationHelper().createFormulaEvaluator();
        // formEval.setIgnoreMissingWorkbooks(true);
        FormulaEvaluator formEval = null;
        
        Sheet sheet = workbook.getSheetAt(0);
        
        convert(sheet, formEval, outFile);
    }

    void convert(Sheet sheet, FormulaEvaluator formEval, Path outFile) throws IOException {
        try (BufferedWriter out = Files.newBufferedWriter(outFile)) {
            convert(sheet, formEval, out);
        }
    }

    void convert(Sheet sheet, FormulaEvaluator formEval, BufferedWriter out) throws IOException {
        
        int maxRow = sheet.getLastRowNum();
        int start = Math.max(0, maxRow-1000);
        
        if (maxRow > 0) return;
        for (int i = start; i<=maxRow; i++) {
            Row row =  sheet.getRow(i);
            List<Object> values;
            if (row != null) {
                values = rowToValues(row, formEval);
            } else {
                values = List.of();
            }
            
            String line = values.stream()
                    .map( v -> v != null ? v.toString() : "")
                    .collect(Collectors.joining(SEP));
            
            out.write(line);
            out.newLine();
        }
    }

    List<Object> rowToValues(Row row, FormulaEvaluator formEval) {
        
        ModernExcelView.CellCaster<Object> caster = new NoFormsCellCaster();
        
        return StreamSupport.stream(row.spliterator(), false)
                .map( cell -> {
                    Object val = caster.cast(cell, formEval);
                    return val;
                })
                .collect(Collectors.toList());
        
    }
    
    protected static class NoFormsCellCaster implements ModernExcelView.CellCaster<Object> {

        @Override
        public Object cast(Cell cell, FormulaEvaluator formEval) {
            if (cell == null) return null;            
            switch(cell.getCellType()) {
                case STRING: return cell.getRichStringCellValue().getString().trim();
                case NUMERIC: return cell.getNumericCellValue();
                case BOOLEAN: return cell.getBooleanCellValue();
                case FORMULA: return null;
                default: return null;
            }            
        }        
    }    
    
}
