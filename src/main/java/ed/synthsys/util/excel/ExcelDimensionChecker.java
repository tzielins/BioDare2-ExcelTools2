/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ed.synthsys.util.excel;

import java.io.IOException;
import java.nio.file.Path;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.NotOLE2FileException;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 *
 * @author Tomasz Zielinski <tomasz.zielinski@ed.ac.uk>
 */
public class ExcelDimensionChecker {
    
    static final XLSXDimensionsChecker XLSX_CHECKER = new XLSXDimensionsChecker();
    static final XLSDimensionsChecker XLS_CHECKER = new XLSDimensionsChecker();
    
    public static int[] rowsColsDimensions(Path file) throws IOException, InvalidFormatException {
        
        if (isOldXLS(file)) {
            return XLS_CHECKER.rowsColsDimensions(file);
        } else {
            return XLSX_CHECKER.rowsColsDimensions(file);
        }
    }
    
    public static boolean isOldXLS(Path file) throws IOException, InvalidFormatException {
        
        try (POIFSFileSystem fs = new POIFSFileSystem(file.toFile(), true)) {
            if (fs != null) return true;
        } catch (OfficeXmlFileException e) { //xlsx format
        } catch (NotOLE2FileException e) {
          throw new InvalidFormatException("Not an excel file",e);
        }
        return false;
    }
}
