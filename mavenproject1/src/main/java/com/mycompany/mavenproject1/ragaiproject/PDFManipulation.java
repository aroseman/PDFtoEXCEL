/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.mavenproject1.ragaiproject;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import static java.util.Collections.list;
import java.util.List;
import java.util.Vector;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.DefaultListModel;
import javax.swing.JList;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentCatalog;
import org.apache.pdfbox.pdmodel.interactive.form.PDAcroForm;
import org.apache.pdfbox.pdmodel.interactive.form.PDField;
import org.apache.pdfbox.pdmodel.interactive.form.PDNonTerminalField;
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
public class PDFManipulation {

    private String[] fieldValues;
    // private Pair[] fieldLabelPairs = new Pair[1000];
    private List<Pair> fieldLabelPairs = new ArrayList<Pair>() ;
    private String[] arrayOfFieldNames = new String[1000];
    private int arrayIndex = 0;

    

    /*  public void powderMETForm(PDDocument pdfDocument, JList list) throws IOException {
        String[] formattedFields = null;
        PDDocumentCatalog docCatalog = pdfDocument.getDocumentCatalog();
        PDAcroForm acroForm = docCatalog.getAcroForm();
        java.util.List<PDField> fields = acroForm.getFields();

        System.out.println(fields.size() + " top-level fields were found on the form");
        for(PDField field : fields){
            this.arrayOfFieldNames[arrayIndex] = field.getPartialName();
            list.setListData(arrayOfFieldNames);
            this.arrayIndex++;
        }

       /* for (PDField field : fields) {
            if (field.getPartialName().equals("Name")) {
                this.name = field.getValueAsString();
            } else if (field.getPartialName().equals("School")) {
                this.school = field.getValueAsString();
            } else if (field.getPartialName().equals("PhD Student") || field.getPartialName().equals("MS Student") || field.getPartialName().equals(" Undergraduate Student")) {
                this.studentStatus.add(field.getValueAsString());
            } else if (field.getPartialName().equals("Email")) {
                this.email = field.getValueAsString();
            } else if (field.getPartialName().equals("Phone")) {
                this.telephone = field.getValueAsString();
            } else if (field.getPartialName().equals("Oral Presentation") || field.getPartialName().equals("Poster Presentation")) {
                this.posterOrPaper.add(field.getValueAsString());
            } else if (field.getPartialName().equals("Paper Title")) {
                this.paperOrPosterTitle = field.getValueAsString();
            } else if (field.getPartialName().equals("CoAuthor") || field.getPartialName().equals("Lead Author")) {
                this.presenterType.add(field.getValueAsString());
            } else if (field.getPartialName().equals("Yes") || field.getPartialName().equals("No")) {
                this.nsfStatus.add(field.getValueAsString());
            } else if (field.getPartialName().equals("Grant")) {
                this.grantNumber = field.getValueAsString();
            } else if (field.getPartialName().equals("Advisor Name")) {
                this.advisor = field.getValueAsString();
            } else if (field.getPartialName().equals("Check here if you are a member of an underrepresented group")) {
                this.minorityStatus = field.getValueAsString();
            } else {

            }
            System.out.println(field.getPartialName() + ": " + field.getValueAsString());*/
    //processField(field, "|--", field.getPartialName());
    //i++;
    //}
    //}

    /* public void msecForm(PDDocument pdfDocument) throws IOException {
        PDDocumentCatalog docCatalog = pdfDocument.getDocumentCatalog();
        PDAcroForm acroForm = docCatalog.getAcroForm();
        java.util.List<PDField> fields = acroForm.getFields();

        System.out.println(fields.size() + " top-level fields were found on the form");

        for (PDField field : fields) {

            if (field.getPartialName() == "Name") {

                this.name = field.getValueAsString();

            } else if (field.getPartialName() == "") {
                processField(field, "|--", field.getPartialName());
            }
        }
    }*/

 /*private void processField(PDField field, String sLevel, String sParent) throws IOException {
        String partialName = field.getPartialName();

        if (field instanceof PDNonTerminalField) {
            if (!sParent.equals(field.getPartialName())) {
                if (partialName != null) {
                    sParent = sParent + "." + partialName;
                }
            }
            System.out.println(sLevel + sParent);

            for (PDField child : ((PDNonTerminalField) field).getChildren()) {
                processField(child, "|  " + sLevel, sParent);
            }
        } else {
            String fieldValue = field.getValueAsString();
            StringBuilder outputString = new StringBuilder(sLevel);
            outputString.append(sParent);
            if (partialName != null) {
                outputString.append(".").append(partialName);
            }
            outputString.append(" = ").append(fieldValue);
            outputString.append(",  type=").append(field.getClass().getName());
            System.out.println(outputString);
        }
    }*/
    public String[] getFieldNames(PDDocument pdfDocument, JList list) {
        int i = 0;
        
        String[] names = new String[10000];
        PDDocumentCatalog docCatalog = pdfDocument.getDocumentCatalog();
        PDAcroForm acroForm = docCatalog.getAcroForm();
        java.util.List<PDField> fields = acroForm.getFields();

        for (PDField field : fields) {
            names[i] = field.getPartialName();
            this.arrayOfFieldNames[i] = names[i];
            list.setListData(names);
            this.fieldLabelPairs.add(new Pair(field.getPartialName(), field.getValueAsString()));
            System.out.println(this.fieldLabelPairs.size());
            i++;

        }
        
        return names;
    }

    public void convert(List<PDDocument> pdfList, List<String> selectedFields, String fileName) {

        
        Workbook wb = new HSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();
        HSSFSheet s1 = (HSSFSheet) wb.createSheet("Sheet 1");

        Row header = s1.createRow((short) 0);
        
        //initialize column headers
        for (int i = 0; i < selectedFields.size(); i++) {
            Cell headerCell = header.createCell(i);
            headerCell.setCellValue(selectedFields.get(i));
        }
        
        //for(int i = 0; i < selectedFields.size();i++){ //fills out row
            //Cell dataCell = data.createCell(i);
            
            for(int y = 0; y < pdfList.size(); y++){
                PDDocumentCatalog docCatalog = pdfList.get(y).getDocumentCatalog();
                PDAcroForm acroForm = docCatalog.getAcroForm();
                java.util.List<PDField> fields = acroForm.getFields();
                Row data = s1.createRow((short) y+1);
                for(int i = 0; i < selectedFields.size();i++){
                    Cell dataCell = data.createCell(i);
                    for (PDField field : fields) {
                    System.out.println("Field Value: " + field.getValueAsString());
                    if(field.getPartialName().equals(selectedFields.get(i))){
                        
                        dataCell.setCellValue(field.getValueAsString());
                    }
                    
            

        }
                }
                
               /* for(int j = 0; j < this.fieldLabelPairs.size();j++){
                if(this.fieldLabelPairs.get(j).getLabel().equals(selectedFields.get(i))){
                    dataCell.setCellValue(this.fieldLabelPairs.get(j).getValue());
                    
                }
            }*/
            }
            
        //}
       /*for (int i = 0; i < selectedFields.size(); i++){
           Row data = s1.createRow(i+1);
           
           for(int j = 0; j< this.fieldLabelPairs.length; j++){
               Cell dataCell  = data.createCell(i);
               if(this.fieldLabelPairs[j].getLabel().equals(selectedFields.get(i))){
                   dataCell.setCellValue(this.fieldLabelPairs[j].getValue());
               }
               
               
       }
       }*/

        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(fileName + ".xls");
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
