/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package ed.synthsys.util.excel;

import java.io.File;
import java.time.LocalDate;
import java.time.Month;
import java.time.temporal.Temporal;
import java.util.Arrays;
import java.util.List;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author tzielins
 */
public class ModernExcelViewTest {
    
    static final double EPS = 1E-6;
    public ModernExcelViewTest() {
    }
    
    ModernExcelView makeInstance(File file) throws Exception {
        return new ModernExcelView(file);
    }


    /**
     * Test of isExcelFile method, of class ModernExcelView.
     */
    @Test
    public void testIsExcelFile() throws Exception {
        File file = new File(getClass().getResource("SimpleImagingData.xlsx").getFile());       
        assertTrue(ModernExcelView.isExcelFile(file));
    }

    /**
     * Test of selectSheet method, of class ModernExcelView.
     */
    @Test
    public void testSelectSheet() throws Exception {
        File file = new File(getClass().getResource("SimpleImagingData.xlsx").getFile()); 
        try (ModernExcelView instance = makeInstance(file)) {
            
            String val = instance.readStringCell(0,0);
            assertNull(val);
            instance.selectSheet(1);
            val = instance.readStringCell(0,0);
            assertEquals("second sheet",val);
        };
    }

    /**
     * Test of readStringRow method, of class ModernExcelView.
     */
    @Test
    public void testReadStringRow() throws Exception {
        File file = new File(getClass().getResource("SimpleImagingData.xlsx").getFile()); 
        try (ModernExcelView instance = makeInstance(file)) {
        
            int rowNr = 1;
            int firstCol = 0;
        
            List<String> expResult = Arrays.asList(null,null,"Id","3.2","4","5","6","7","8","9","10","11","12","13");
            
            List<String> result = instance.readStringRow(rowNr, firstCol);
            assertEquals(expResult, result);
            
            rowNr = 1;
            firstCol = 3;
            
            expResult = Arrays.asList("3.2","4","5","6","7","8","9","10","11","12","13");
            
            result = instance.readStringRow(rowNr, firstCol);
            assertEquals(expResult, result);  
            
            rowNr = 1;
            firstCol = 3;
            int lastCol = 4;
            expResult = Arrays.asList("3.2","4");
            
            result = instance.readStringRow(rowNr, firstCol,lastCol);
            assertEquals(expResult, result);  
        }
    }


    /**
     * Test of getLastColumn method, of class ModernExcelView.
     */
    @Test
    public void testGetLastColumn() throws Exception {
        File file = new File(getClass().getResource("SimpleImagingData.xlsx").getFile()); 
        try (ModernExcelView instance = makeInstance(file)) {
            instance.selectSheet(1);
            int rowNr = 0;
            int expResult = 0;
            int result = instance.getLastColumn(rowNr);
            assertEquals(expResult, result);
            
            rowNr = 2;
            expResult = 5;
            
            result = instance.getLastColumn(rowNr);
            assertEquals(expResult, result);            
        }
    }

    /**
     * Test of readDoubleColumn method, of class ModernExcelView.
     */
    @Test
    public void testReadDoubleColumn() throws Exception {
        File file = new File(getClass().getResource("SimpleImagingData.xlsx").getFile()); 
        try (ModernExcelView instance = makeInstance(file)) {
        
            int colNr = 0;
            int firstRow = 10;
            List<Double> expResult = Arrays.asList(null,null,null);
            List<Double> result = instance.readDoubleColumn(colNr, firstRow);
            assertEquals(expResult, result);
            
            colNr = 1;
            firstRow = 3;
            expResult = Arrays.asList(1.0,2.0,3.0,4.0,5.0,6.0,7.0,8.0,9.0,10.0);
            result = instance.readDoubleColumn(colNr, firstRow);
            assertEquals(expResult, result);            
            
            colNr = 13;
            firstRow = 2;
            int lastRow = 5;
            expResult = Arrays.asList(null,881409.5,971627.5,887073.0);
            
            result = instance.readDoubleColumn(colNr, firstRow,lastRow);
            assertEquals(expResult, result);            
            
        }
    }
    
    @Test
    public void testReadColumns() throws Exception {
        File file = new File(getClass().getResource("SimpleImagingData.xlsx").getFile()); 
        try (ModernExcelView instance = makeInstance(file)) {
        
            int colNr = 1;
            int lastColNr = 3;
            int firstRow = 3;
            int lastRow = 13;
            
            ModernExcelView.CellCaster<Double> caster = new ModernExcelView.DoubleCellCaster();
            
            List<List<Double>> result = instance.readColumns(colNr, lastColNr, firstRow, lastRow, caster);
            
            assertEquals(lastColNr-colNr+1,result.size());
            for (List<Double> list : result)
                assertEquals(lastRow-firstRow+1, list.size());
            
            assertEquals(1,result.get(0).get(0),EPS);
            assertEquals(10,result.get(0).get(9),EPS);
            assertNull(result.get(0).get(10));
            
            assertNull(result.get(1).get(0));
            assertNull(result.get(1).get(9));
            assertNull(result.get(1).get(10));

            assertEquals(1487.95276565749,result.get(2).get(0),EPS);
            assertEquals(1480.07583682999,result.get(2).get(9),EPS);
            assertNull(result.get(2).get(10));

        }
    }
    
    @Test
    public void testFindParam() throws Exception {
        File file = new File(getClass().getResource("SimpleParamData.xlsx").getFile()); 
        try (ModernExcelView instance = makeInstance(file)) {
        
            String pName = "name";
            String pVal = "Tomasz";
            String res = instance.findParam(pName, 0, 10, ModernExcelView.STRING_CASTER);
            
            assertEquals(pVal, res);
            
            pName = "age";
            pVal = null;
            res = instance.findParam(pName, 0, 10, ModernExcelView.STRING_CASTER);
            assertEquals(pVal, res);           
            
            pName = "comment";
            pVal = "First comment";
            res = instance.findParam(pName, 0, 10, ModernExcelView.STRING_CASTER);
            assertEquals(pVal, res);           
            
            pName = "comment";
            pVal = "Ignored coment";
            res = instance.findParam(pName, 5, 10, ModernExcelView.STRING_CASTER);
            assertEquals(pVal, res);           
            
            pName = "year";
            pVal = "2015";
            res = instance.findParam(pName, 0, 11, ModernExcelView.STRING_CASTER);
            assertEquals(pVal, res);           
            
            pName = "year";
            pVal = null;
            res = instance.findParam(pName, 0, 8, ModernExcelView.STRING_CASTER);
            assertEquals(pVal, res);           
        }
    }
    
    @Test
    public void testGetSheetName() throws Exception {
        
        File file = new File(getClass().getResource("SimpleImagingData.xlsx").getFile()); 
        try (ModernExcelView instance = makeInstance(file)) {

            String exp = "Sheet1";
            String res = instance.getCurrentSheetName();
            assertEquals(exp, res);
            
            instance.selectSheet(2);
            exp = "Third";
            res = instance.getCurrentSheetName();
            assertEquals(exp, res);            
        }
    }
    
    @Test
    public void testGetSheetNr() throws Exception {
        
        File file = new File(getClass().getResource("SimpleImagingData.xlsx").getFile()); 
        try (ModernExcelView instance = makeInstance(file)) {

            int exp = 0;
            int res = instance.getCurrentSheetNr();
            assertEquals(exp, res);
            
            instance.selectSheet(2);
            exp = 2;
            res = instance.getCurrentSheetNr();
            assertEquals(exp, res);            
        }
    }
    
    @Test
    public void testGetTemporalCell() throws Exception {
        
        File file = new File(getClass().getResource("SimpleImagingData.xlsx").getFile()); 
        try (ModernExcelView instance = makeInstance(file)) {

            instance.selectSheet(1);
            int row = 3;
            int col = 2;
            
            LocalDate exp = LocalDate.of(2015, Month.SEPTEMBER, 1);
            LocalDate res = LocalDate.from(instance.readTemporalCell(row, col));
            
            assertEquals(exp,res);
        }
    }
    
    @Test
    public void testGetDoubleCell() throws Exception {
        
        File file = new File(getClass().getResource("SimpleImagingData.xlsx").getFile()); 
        try (ModernExcelView instance = makeInstance(file)) {

            int row = 3;
            int col = 3;
            
            double exp = 1487.95276565749;
            double res = instance.readDoubleCell(row, col);
            
            assertEquals(exp,res,1E-6);
        }
    }
    

}