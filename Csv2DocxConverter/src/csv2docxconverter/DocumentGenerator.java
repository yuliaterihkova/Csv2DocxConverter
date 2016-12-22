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
     * @param columns a list of columns for data parsing
     * @param content list of data rows for parsing
     * @return an XWPFDocument representing a DocX file
    */        
    public XWPFDocument generateDocx(String[] columns, List content){
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
        setHeader(table, columns);
        
        // if account list is empty
        if(content == null){
            // add empty row to the table
            XWPFTableRow emptyRow = table.createRow();
        }else{
            // create rows in table
            for(Object object : content) {
                String[] csvRow = (String[]) object;
                
                XWPFTableRow row = table.createRow();
                //create cells in a row
                for(int i = 0; i < csvRow.length; i++){
                    XWPFTableCell cell = row.getCell(i);
                    XWPFRun run = setBodyCell(cell);
                    run.setText(csvRow[i]);
                }
            }
        }
        
        return document;
    }
     
    /**
    * Create header for a table
    */  
    private static void setHeader(XWPFTable table, String[] columns){
        // set initial cell
        XWPFTableRow tableRowOne = table.getRow(0);    
        XWPFTableCell cell = tableRowOne.getCell(0);
        XWPFRun run = setHeaderCell(cell);
        run.setText(columns[0]);
        
        // add other cells
        for(int i = 1; i < columns.length; i++){
            cell = tableRowOne.addNewTableCell();
            run = setHeaderCell(cell);
            run.setText(columns[i]);
        }
    }
    
    /**
    * Set header cell style
    */ 
    private static XWPFRun setHeaderCell(XWPFTableCell cell){
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
    private static XWPFRun setBodyCell(XWPFTableCell cell){
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        
        XWPFParagraph paragraph = cell.getParagraphArray(0);
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        paragraph.setSpacingBefore(4);
        paragraph.setSpacingAfter(4);
        
        return paragraph.createRun();
    }
}
