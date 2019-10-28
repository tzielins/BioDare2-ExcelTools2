/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ed.synthsys.util.excel;

import static ed.synthsys.util.excel.ExcelDimensionChecker.isOldXLS;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 *
 * @author Tomasz Zielinski <tomasz.zielinski@ed.ac.uk>
 */
public class Excel2TextConverter {

    final XLSX2TextConverter xlsxConverter = new XLSX2TextConverter();
    final Workbook2TextConverter workbookConverter = new Workbook2TextConverter();

    int WORKBOOK_SIZE_THRESHOLD = 25 * 1024 * 1024; // 25Mb

    public void convert(Path inFile, Path outFile) throws IOException, InvalidFormatException {
        convert(inFile, outFile, ",");
    }

    public void convert(Path inFile, Path outFile, String SEP) throws IOException, InvalidFormatException {

        if (Files.size(inFile) < WORKBOOK_SIZE_THRESHOLD) {
            workbookConverter.convert(inFile, outFile, SEP);
        } else {
            if (isOldXLS(inFile)) {
                throw new IOException("Cannot convert old xls files larger than "
                        + WORKBOOK_SIZE_THRESHOLD + " please use xlsx format instead");
            } else {
                xlsxConverter.convert(inFile, outFile, SEP);
            }
        }
    }

}
