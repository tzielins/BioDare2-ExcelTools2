/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ed.synthsys.util.excel;

import java.nio.file.Path;
import java.nio.file.Paths;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Before;
import org.junit.Test;
import static org.junit.Assert.*;
import org.junit.Ignore;

/**
 *
 * @author Tomasz Zielinski <tomasz.zielinski@ed.ac.uk>
 */
public class ExcelDimensionCheckerTest {
    
    public ExcelDimensionCheckerTest() {
    }
    
    @Before
    public void setUp() {
    }

    @Test
    public void testChecksFormats() throws Exception {
        
        Path file = Paths.get(this.getClass().getResource("SimpleImagingData.xlsx").toURI());
        assertFalse(ExcelDimensionChecker.isOldXLS(file));
        
        file = Paths.get(this.getClass().getResource("SimpleImagingData.xls").toURI());
        assertTrue(ExcelDimensionChecker.isOldXLS(file));
        
        try {
            file = Paths.get(this.getClass().getResource("2CSVTest.csv").toURI());
            assertFalse(ExcelDimensionChecker.isOldXLS(file));
            fail("Exception expected");            
        } catch (InvalidFormatException e) {}
    }

    @Test
    public void testChecksSizes() throws Exception {
        
        int[] exp = {13, 14};
        
        
        Path file = Paths.get(this.getClass().getResource("SimpleImagingData.xlsx").toURI());        
        int[] rowsCols = ExcelDimensionChecker.rowsColsDimensions(file);        
        assertArrayEquals(exp, rowsCols);        
        
        file = Paths.get(this.getClass().getResource("SimpleImagingData.xls").toURI());
        rowsCols = ExcelDimensionChecker.rowsColsDimensions(file);        
        assertArrayEquals(exp, rowsCols);        
        
        try {
            file = Paths.get(this.getClass().getResource("2CSVTest.csv").toURI());
            rowsCols = ExcelDimensionChecker.rowsColsDimensions(file);     
            fail("Exception expected");            
        } catch (InvalidFormatException e) {}
    }    
    
    @Test
    @Ignore("Test file not commited")
    public void testCanCheckLargeFiles() throws Exception {
        
        
        
        
        Path file = Paths.get("E:\\Temp\\long_10000x1200.xlsx"); 
        int[] rowsCols = ExcelDimensionChecker.rowsColsDimensions(file);        
        int[] exp = {1202, 10001};
        assertArrayEquals(exp, rowsCols);        
        
        file = Paths.get("E:\\Temp\\long_255x10000.xls");
        rowsCols = ExcelDimensionChecker.rowsColsDimensions(file);        
        exp = new int[]{10001,256};    
        assertArrayEquals(exp, rowsCols);        
        
    }     
}
