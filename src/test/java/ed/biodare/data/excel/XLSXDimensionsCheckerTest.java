/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ed.biodare.data.excel;

import java.nio.file.Path;
import java.nio.file.Paths;
import org.junit.Before;
import org.junit.Test;
import static org.junit.Assert.*;
import org.junit.Ignore;

/**
 *
 * @author Tomasz Zielinski <tomasz.zielinski@ed.ac.uk>
 */
public class XLSXDimensionsCheckerTest {
    
    public XLSXDimensionsCheckerTest() {
    }
    
    XLSXDimensionsChecker instance;
    
    @Before
    public void setUp() {
        
        instance = new XLSXDimensionsChecker();
    }

    @Test
    public void testChecksDimenssions() throws Exception {
        
        Path inFile = Paths.get(this.getClass().getResource("2CSVTest.xlsx").toURI());
        
        int[] rowsCols = instance.rowsColsDimensions(inFile);

        int[] exp = {6, 5};
        

        assertArrayEquals(exp, rowsCols);
            
    }  
    
    @Test
    public void testChecksDimenssionsFromTheFirstSheet() throws Exception {
        
        Path inFile = Paths.get(this.getClass().getResource("SimpleImagingData.xlsx").toURI());
        
        int[] rowsCols = instance.rowsColsDimensions(inFile);

        int[] exp = {13, 14};
        

        assertArrayEquals(exp, rowsCols);
            
    }  
    
    
    
    @Test
    @Ignore("Not commited test file")
    public void testCanCheckLargeFiles() throws Exception {
        
        Path inFile = Paths.get("E:\\Temp\\long_10000x1200.xlsx");
        
        int[] rowsCols = instance.rowsColsDimensions(inFile);

        int[] exp = {1202, 10001};
        

        assertArrayEquals(exp, rowsCols);
    }    
    
   
    
}
