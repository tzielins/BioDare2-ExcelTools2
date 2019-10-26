/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ed.synthsys.util.excel;

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import org.junit.Before;
import org.junit.Test;
import static org.junit.Assert.*;
import org.junit.Ignore;
import org.junit.Rule;
import org.junit.rules.TemporaryFolder;


/**
 *
 * @author Tomasz Zielinski <tomasz.zielinski@ed.ac.uk>
 */
public class XLSX2TextConverterTest {
    
    @Rule
    public TemporaryFolder testFolder = new TemporaryFolder();
    
    public XLSX2TextConverterTest() {
    }
    
    @Before
    public void setUp() {
    }

    @Test
    @Ignore("Not commited test file")
    public void testCanSaveLargeToCSV() throws Exception {
        
        Path inFile = Paths.get("E:\\Temp\\long_5000x1200.xlsx");
        Path outFile = inFile.getParent().resolve(inFile.getFileName().toString()+".xlsxxls2csv.csv");   
        
        XLSX2TextConverter xlsx2csv = new XLSX2TextConverter();
        xlsx2csv.convert(inFile, outFile);
        
        assertTrue(Files.isRegularFile(outFile));
        assertTrue(Files.size(outFile) > 1000);
        
        assertEquals(1201, countLines(outFile));
        assertEquals(5001, countColumns(outFile));
    }
    
    int countColumns(Path file) throws IOException {
        try (BufferedReader in = Files.newBufferedReader(file)) {
            
            String line = in.readLine();
            if (line != null) {
                return line.split(",").length;
            } else {
                return 0;
            }
        }
    }
    
    int countLines(Path file) throws IOException {
        int lines = 0;
        try (BufferedReader in = Files.newBufferedReader(file)) {
            for (;;) {
                String line = in.readLine();
                if (line == null) break;
                lines++;
            }
        }
        return lines;
    }
    
    @Test
    public void testConvertsTestFileCSV() throws Exception {
        
            XLSX2TextConverter xlsx2csv = new XLSX2TextConverter();
            
            Path inFile = Paths.get(this.getClass().getResource("2CSVTest.xlsx").toURI());
            Path outFile = testFolder.newFile().toPath();   
            xlsx2csv.convert(inFile, outFile);
            
            assertTrue(Files.isRegularFile(outFile));
            assertTrue(Files.size(outFile) > 10);
            
            List<String> lines = Files.readAllLines(outFile);
            
            // lines.forEach( System.out::println);
            
            List<String> exp = List.of(
                    "",
                    "A,B,C",                    
                    "",
                    ",1.1,1.2,2.3434432423434199E18",
                    "1.0,2.0",
                    ",,A,B,\"with,comma\""
            );
            
            assertEquals(exp, lines);
            
    }    
}
