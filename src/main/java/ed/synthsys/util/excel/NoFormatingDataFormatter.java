/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.

   2012 - Alfresco Software, Ltd.
   Alfresco Software has modified source of this file
   The details of changes as svn diff can be found in svn at location root/projects/3rd-party/src 
==================================================================== */
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
