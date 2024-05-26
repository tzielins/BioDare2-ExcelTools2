/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ed.biodare.data.excel;

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
public class Excel2TextConverterTest {
    
    @Rule
    public TemporaryFolder testFolder = new TemporaryFolder();
    
    public Excel2TextConverterTest() {
    }
    
    Excel2TextConverter instance;
    
    @Before
    public void setUp() {
        instance = new Excel2TextConverter();
    }

    @Test
    @Ignore("Not commited the large tets file as >156Mg and gitub refuses")
    public void testCanSaveLargeToCSV() throws Exception {
        
        Path inFile = Paths.get(this.getClass().getResource("long_10000x1200.xlsx").toURI());
        Path outFile = inFile.getParent().resolve(inFile.getFileName().toString()+".excel2text.csv");   
        
        instance.convert(inFile, outFile);
        
        
        assertTrue(Files.isRegularFile(outFile));
        assertTrue(Files.size(outFile) > 10000*1200);
    }
    
    @Test
    public void testConvertsXLSXFile() throws Exception {
        
            
            Path inFile = Paths.get(this.getClass().getResource("2CSVTest.xlsx").toURI());
            Path outFile = testFolder.newFile().toPath();   
            instance.convert(inFile, outFile);
            
            assertTrue(Files.isRegularFile(outFile));
            assertTrue(Files.size(outFile) > 10);
            
            List<String> lines = Files.readAllLines(outFile);
            
            //lines.forEach( System.out::println);
            
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
    
    @Test
    public void testConvertsXLSFile() throws Exception {
        
            
            Path inFile = Paths.get(this.getClass().getResource("2CSVTest.xls").toURI());
            Path outFile = testFolder.newFile().toPath();   
            instance.convert(inFile, outFile);
            
            assertTrue(Files.isRegularFile(outFile));
            assertTrue(Files.size(outFile) > 10);
            
            List<String> lines = Files.readAllLines(outFile);
            
            //lines.forEach( System.out::println);
            
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
    
    @Test
    //@Ignore("Not commited test file")
    public void testConvertsMediumXLSFile() throws Exception {
        
            
            Path inFile = Paths.get(this.getClass().getResource("long_255x5000.xls").toURI());
            Path outFile = testFolder.newFile().toPath();   
            instance.convert(inFile, outFile);
            
            assertTrue(Files.isRegularFile(outFile));
            assertTrue(Files.size(outFile) > 5000);
            
            
    }      
    
}
