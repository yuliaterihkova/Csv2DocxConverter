/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package csv2docxconverter;

import java.math.BigInteger;
import java.util.List;
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
     * @param columnNumbers a list of columns for data parsing
     * @param content list of data rows for parsing
     * @return an XWPFDocument representing a DocX file
    */        
    public XWPFDocument generateDocx(int[] columnNumbers, List content){
        XWPFDocument document = new XWPFDocument();

        // create title
        XWPFParagraph title = document.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleRun = title.createRun();
        titleRun.setText("G SUITE created accounts");
        titleRun.setFontSize(18); 
        
        // create account table
        XWPFTable table = document.createTable();
        // set "justified" alignment
        table.getCTTbl().addNewTblPr().addNewTblW().setW(BigInteger.valueOf(10000));
        
        //create header for table
        setHeader(table, columnNumbers);
        
        // if account list is empty
        if(content == null || content.size() == 0){
            // add empty row to the table
            XWPFTableRow emptyRow = table.createRow();
        }else{
            // create rows in table
            for(Object object : content) {
                String[] csvRow = (String[]) object;
                
                XWPFTableRow row = table.createRow();
                //create cells in a row
                for(int i = 0; i < columnNumbers.length; i++){
                    int number = columnNumbers[i];
                    
                    XWPFTableCell cell = row.getCell(i);
                    if(cell != null){
                        XWPFRun run = setBodyCell(cell);
                        if(number - 1 <= csvRow.length){
                            run.setText(csvRow[number - 1]);
                        }
                    }
                }
            }
        }
        
        return document;
    }
     
    /**
    * Create header for a table
    */  
    private void setHeader(XWPFTable table, int[] columnNumbers){
        // set initial cell
        XWPFTableRow tableRowOne = table.getRow(0);    

        // add other cells
        for(int i = 0; i < columnNumbers.length; i++){
            int number = columnNumbers[i];
             
            XWPFTableCell cell;
            if(i == 0){
                cell = tableRowOne.getCell(0);
            }else{
                cell = tableRowOne.addNewTableCell();
            }
            XWPFRun run = setHeaderCell(cell);
            run.setText(Integer.toString(number));
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
