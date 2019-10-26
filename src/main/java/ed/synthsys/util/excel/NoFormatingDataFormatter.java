package ed.synthsys.util.excel;

import org.apache.poi.ss.formula.ConditionalFormattingEvaluator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;



class NoFormatingDataFormatter extends DataFormatter {


    @Override
    public String formatRawCellContents(double value, int formatIndex, String formatString, boolean use1904Windowing) {
        
        
        return String.valueOf(value);
    }

    String getUnFormattedNumberString(Cell cell, ConditionalFormattingEvaluator cfEvaluator) {
        if (cell == null) {
            return null;
        }
        double d = cell.getNumericCellValue();
        return String.valueOf(d);
    }

    String getFormattedDateString(Cell cell, ConditionalFormattingEvaluator cfEvaluator) {
        if (cell == null) {
            return null;
        }
        
        return String.valueOf(cell.getNumericCellValue());
    }


    @Override
    public String formatCellValue(Cell cell, FormulaEvaluator evaluator, ConditionalFormattingEvaluator cfEvaluator) {
        
        if (cell == null) {
            return "";
        }

        CellType cellType = cell.getCellType();
        if (cellType == CellType.FORMULA) {
            if (evaluator == null) {
                return cell.getCellFormula();
            }
            cellType = evaluator.evaluateFormulaCell(cell);
        }
        switch (cellType) {
            case NUMERIC :

                if (DateUtil.isCellDateFormatted(cell, cfEvaluator)) {
                    return getFormattedDateString(cell, cfEvaluator);
                }
                return getUnFormattedNumberString(cell, cfEvaluator);

            case STRING :
                return cell.getRichStringCellValue().getString();

            case BOOLEAN :
                return cell.getBooleanCellValue() ? "TRUE" : "FALSE";
            case BLANK :
                return "";
            case ERROR:
                return FormulaError.forInt(cell.getErrorCellValue()).getString();
            default:
                throw new RuntimeException("Unexpected celltype (" + cellType + ")");
        }
    }




    
}
