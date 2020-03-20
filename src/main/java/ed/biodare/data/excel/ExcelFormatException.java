/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ed.biodare.data.excel;

/**
 *
 * @author tzielins
 */
public class ExcelFormatException extends Exception {
    
    public ExcelFormatException(String msg) {
        super(msg);
    }
    
    public ExcelFormatException(String msg,Throwable err) {
        super(msg,err);
    }    
    
}
