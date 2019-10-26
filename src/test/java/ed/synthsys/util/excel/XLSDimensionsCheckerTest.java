/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ed.synthsys.util.excel;

import java.nio.file.Path;
import java.nio.file.Paths;
import org.junit.Before;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author Tomasz Zielinski <tomasz.zielinski@ed.ac.uk>
 */
public class XLSDimensionsCheckerTest {
    
    public XLSDimensionsCheckerTest() {
    }
    
    XLSDimensionsChecker instance;
    
    @Before
    public void setUp() {
        instance = new XLSDimensionsChecker();
    }

    @Test
    public void testChecksDimenssionsFromTheFirstSheet() throws Exception {
        
        Path inFile = Paths.get(this.getClass().getResource("SimpleImagingData.xls").toURI());
        
        int[] rowsCols = instance.rowsColsDimensions(inFile);

        int[] exp = {13, 14};
        

        assertArrayEquals(exp, rowsCols);
            
    } 
    
    @Test
    //@Ignore("Not commited test file")
    public void testCanCheckLargeXlSFile() throws Exception {
        
        Path inFile = Paths.get("E:\\Temp\\long_255x10000.xls");
        
        int[] rowsCols = instance.rowsColsDimensions(inFile);

        int[] exp = {10001,256};
        

        assertArrayEquals(exp, rowsCols);
    }     
    
}
