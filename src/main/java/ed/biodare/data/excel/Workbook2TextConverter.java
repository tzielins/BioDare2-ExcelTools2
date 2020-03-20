/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ed.biodare.data.excel;

import ed.biodare.data.excel.ModernExcelView.CellCaster;
import java.io.BufferedWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;
import java.util.stream.Collectors;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author Tomasz Zielinski <tomasz.zielinski@ed.ac.uk>
 */
class Workbook2TextConverter {
    
    public void convert(Path inFile, Path outFile) throws IOException {
        convert(inFile, outFile, ",");
    }
    
    public void convert(Path inFile, Path outFile, String SEP) throws IOException {
        
        try (Workbook workbook = WorkbookFactory.create(inFile.toFile())) {
            if (workbook.getNumberOfSheets() < 1) {
                throw new IOException("No sheets in file");
            }

            FormulaEvaluator formEval = workbook.getCreationHelper().createFormulaEvaluator();
            formEval.setIgnoreMissingWorkbooks(true); 

            Sheet sheet = workbook.getSheetAt(0);
            convert(sheet, formEval, outFile, SEP);
        };
    }    
    
    void convert(Sheet sheet, FormulaEvaluator formEval, Path outFile, String SEP) throws IOException {
        try (BufferedWriter out = Files.newBufferedWriter(outFile)) {
            convert(sheet, formEval, out, SEP);
        }
    }

    void convert(Sheet sheet, FormulaEvaluator formEval, BufferedWriter out, String SEP) throws IOException {
        
        final CellCaster<Object> caster = ModernExcelView.NATURAL_CASTER;
        Function<Cell, Object> valueExtractor = cell -> caster.cast(cell, formEval);
        
        
        int maxRow = sheet.getLastRowNum();
        
        for (int i = 0; i<=maxRow; i++) {
            Row row =  sheet.getRow(i);
            List<Object> values;
            if (row != null) {
                values = rowToValues(row, valueExtractor);
            } else {
                values = List.of();
            }
            
            String line = values.stream()
                    .map( v -> v != null ? v.toString() : "")
                    .map( v -> {
                        if (v.contains(SEP)) return "\""+v+"\"";
                        else return v;
                    })
                    .collect(Collectors.joining(SEP));
            
            out.write(line);
            out.newLine();
        }
    }

    List<Object> rowToValues(Row row, Function<Cell, Object> valueExtractor) {
        
        int last = row.getLastCellNum();
        List<Object> cells = new ArrayList<>(last);
        
        for (int i = 0; i< last; i++) {
            Cell cell = row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            cells.add(valueExtractor.apply(cell));
        }
        
        return cells;
    }
    
}
