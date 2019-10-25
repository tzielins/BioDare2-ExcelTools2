/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ed.synthsys.util.excel;

import java.io.PrintStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.junit.Before;
import org.junit.Test;
import static org.junit.Assert.*;
import org.junit.Ignore;

/**
 *
 * @author Tomasz Zielinski <tomasz.zielinski@ed.ac.uk>
 */
public class XLS2CSVmraTest {
    
    public XLS2CSVmraTest() {
    }
    
    @Before
    public void setUp() {
    }

    @Test
    @Ignore("The it only works with the old format which cannot have so many columns")
    public void testCanSaveLargeToCSV() throws Exception {
        
        Path inFile = Paths.get("E:\\Temp\\long_10000x1200.xlsx");
        Path outFile = inFile.getParent().resolve(inFile.getFileName().toString()+".mraxls2csv.csv");   
        
        try (PrintStream out = new PrintStream(Files.newOutputStream(outFile))) {
            try (POIFSFileSystem in = new POIFSFileSystem(inFile.toFile(), true)) {
                
                XLS2CSVmra instance = new XLS2CSVmra(in, out, -1);
                instance.process();
                
            }
        }
        assertTrue(Files.isRegularFile(outFile));
        assertTrue(Files.size(outFile) > 1000);
    }
    
}
