/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.mavenproject1.ragaiproject;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author helios
 */
public class ExcelExport {
    String[] columnHeaderArray;
    String[] fieldValueArray;
    PDFManipulation pdf;
    
    public ExcelExport(PDFManipulation pdf){
        this.pdf = pdf;
    }
    
    public void writeSheet() {
        Workbook wb = new HSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();
        HSSFSheet s1 =  (HSSFSheet) wb.createSheet("Sheet 1");
        
        Row row = s1.createRow((short)0);
        Cell cell = row.createCell(0);
        cell.setCellValue(1);
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream("workbook.xls");
            try {
                wb.write(fileOut);
                fileOut.close();
            } catch (IOException ex) {
                Logger.getLogger(ExcelExport.class.getName()).log(Level.SEVERE, null, ex);
            }
            
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelExport.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }

}
