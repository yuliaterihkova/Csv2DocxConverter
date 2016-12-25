/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package csv2docxconverter;

import java.math.BigInteger;
import java.nio.charset.Charset;
import java.util.List;
import java.util.Objects;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 * Class encapsulating DocX file generation logic
 * @author Yulia Terikhova
 */
public class DocumentGenerator {
    
    /**
    * Generate DocX element from list of the rows containing account information
     * @param columnNames a list of columns for data parsing
     * @param contents list of data rows for parsing
     * @return an XWPFDocument representing a DocX file
    */        
    public XWPFDocument generateDocx(String[] columnNames, List contents){
        XWPFDocument document = new XWPFDocument();

        // create title
        XWPFParagraph title = document.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleRun = title.createRun();
        titleRun.setText("G SUITE created accounts");
        titleRun.setFontSize(18); 
       
        for(int k = 0; k < contents.size(); k++) {
            List tableContent = (List)contents.get(k);
            
            // create title
            title = document.createParagraph();
            title.setAlignment(ParagraphAlignment.CENTER);
            titleRun = title.createRun();
            titleRun.setText("Table " + (k + 1));
            titleRun.setFontSize(18);

             // create account table
            XWPFTable table = document.createTable();
            // set "justified" alignment
            table.getCTTbl().addNewTblPr().addNewTblW().setW(BigInteger.valueOf(10000));
       
            //create header for table
            setHeader(table, columnNames);
         
            // if account list is empty
            if(tableContent == null || tableContent.size() == 0){
                // add empty row to the table
                XWPFTableRow emptyRow = table.createRow();
            }else{
                String[] headerRow = null;
                // create rows in table
                for(int i = 0; i < tableContent.size(); i++){  
                    if(i == 0){
                        headerRow = (String[]) tableContent.get(i);
                        continue;
                    }
                    String[] csvRow = (String[]) tableContent.get(i);

                    XWPFTableRow row = table.createRow();
                    //create cells in a row
                    for(int j = 0; j < columnNames.length; j++){
                        int number = getColumnNumber(columnNames[j], headerRow);

                        XWPFTableCell cell = row.getCell(j);
                        if(cell != null){
                            XWPFRun run = setBodyCell(cell);
                            if(number >= 0 && number < csvRow.length){
                                run.setText(csvRow[number]);
                            }
                        }
                    }
                }
            }
        }
        
        return document;
    }
    
    /**
     * Get number of column from column list
     * @param name name for look
     * @param columnNames columns list
     */
    private int getColumnNumber(String name, String[] columnNames){
        for(int i = 0; i < columnNames.length; i++){  
            String name2 = columnNames[i].trim().toLowerCase();
            if(name.contentEquals(name2)){
                return i;
            }
        }
        return -1;
    }
    
    /**
    * Create header for a table
    */  
    private void setHeader(XWPFTable table, String[] columnNames){
        // set initial cell
        XWPFTableRow tableRowOne = table.getRow(0);    

        // add other cells
        for(int i = 0; i < columnNames.length; i++){             
            XWPFTableCell cell;
            if(i == 0){
                cell = tableRowOne.getCell(0);
            }else{
                cell = tableRowOne.addNewTableCell();
            }
            XWPFRun run = setHeaderCell(cell);
            run.setText(columnNames[i]);
        }
    }
    
    /**
    * Set header cell style
    */ 
    private XWPFRun setHeaderCell(XWPFTableCell cell){
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        
        XWPFParagraph paragraph = cell.getParagraphArray(0);
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        paragraph.setSpacingBefore(8);
        paragraph.setSpacingAfter(8);
        
        XWPFRun run = paragraph.createRun();
        run.setBold(true);
        
        return run;
    }
    
    /**
    * Set body cell style
    */ 
    private XWPFRun setBodyCell(XWPFTableCell cell){
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        
        XWPFParagraph paragraph = cell.getParagraphArray(0);
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        paragraph.setSpacingBefore(4);
        paragraph.setSpacingAfter(4);
        
        return paragraph.createRun();
    }
}