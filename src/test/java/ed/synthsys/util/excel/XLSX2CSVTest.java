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
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.junit.Before;
import org.junit.Test;
import static org.junit.Assert.*;
import org.junit.Ignore;

/**
 *
 * @author Tomasz Zielinski <tomasz.zielinski@ed.ac.uk>
 */
public class XLSX2CSVTest {
    
    public XLSX2CSVTest() {
    }
    
    @Before
    public void setUp() {
    }

    @Test
    @Ignore("Not commited test file")
    public void testCanSaveLargeToCSV() throws Exception {
        
        Path inFile = Paths.get("E:\\Temp\\long_10000x1200.xlsx");
        Path outFile = inFile.getParent().resolve(inFile.getFileName().toString()+".xlsxxls2csv.csv");   
        
        try (PrintStream out = new PrintStream(Files.newOutputStream(outFile))) {
        try (OPCPackage p = OPCPackage.open(inFile.toString(), PackageAccess.READ)) {
            XLSX2CSV xlsx2csv = new XLSX2CSV(p, out, -1);
            xlsx2csv.process();
        }
        }
        assertTrue(Files.isRegularFile(outFile));
        assertTrue(Files.size(outFile) > 1000);
    }
}
