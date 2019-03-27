package data;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

public class Debit108 
{
    private final String From;
    private final String To;
    private final String Date;
    private final String Filename;
    private final ArrayList<Object[]> charge_materials;
    private XWPFDocument doc = null;
    
    public Debit108(String From, String To, String Date, String Filename, ArrayList<Object[]> charge_materials) 
    {
        this.From = From;
        this.To = To;
        this.Date = Date;
        this.Filename = Filename;
        this.charge_materials = charge_materials;
        
        try
        {            
            doc = new XWPFDocument(getClass().getClassLoader().getResourceAsStream("rsc/108.docx"));
        }
        catch(Exception e)
        {
            System.out.println("Error creating 108 docx.");
        }
    }
    
    private void writeField(String field, int pos)
    {
        XWPFTable table = doc.getTableArray(pos);
        
        XWPFParagraph paragraph = doc.createParagraph();
        XWPFRun paragraphRun = paragraph.createRun();
        
        paragraphRun.setFontSize(12);
        paragraphRun.setFontFamily("Arial");
        paragraphRun.setText(field);
        
        XWPFTableRow tableRow = table.getRow(0);
        tableRow.getCell(0).setParagraph(paragraph);
        paragraph.removeRun(0);
    }
    
    private void writeDate()
    {
        XWPFTable table = doc.getTableArray(2);
        
        XWPFParagraph paragraph = doc.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        
        XWPFRun paragraphRun = paragraph.createRun();

        paragraphRun.setFontSize(6);
        paragraphRun.setFontFamily("Calibri");
        paragraphRun.setText(Date);
        
        XWPFTableRow tableRow = table.getRow(3);
        tableRow.getCell(6).setParagraph(paragraph);
        paragraph.removeRun(0);
    }
    
    private void writeTable()
    {
        /*
        //Debugging
        String ilika[][] = new String [][] {  {"test","test","test","test","test", "test"},
                                              {"test","test","test","test","test", "test"},
                                              {"test","test","test","test","test", "test"},
                                              {"test","test","test","test","test", "test"},
                                              {"test","test","test","test","test", "test"},
                                              {"test","test","test","test","test", "test"},
                                              {"test","test","test","test","test", "test"},
                                              {"test","test","test","test","test", "test"},
                                              {"test","test","test","test","test", "test"},
                                              {"test","test","test","test","test", "test"},
                                              {"test","test","test","test","test", "test"},
                                              {"1","1","1","1","1", "1"},
                                              {"5","5","5","5","5", "5"},
                                              {"test","test","test","test","test", "test"},
                                              {"test","test","test","test","test", "test"},
                                              {"2","2","2","2","2", "2"},
                                              {"test","test","test","test","test", "test"},
                                              {"2","2","2","2","2", "2"},
                                              {"test","test","test","test","test", "test"},
                                              {"2","2","2","2","2", "2"},
                                              {"test","test","test","test","test", "test"}
                                           };                                     
        */
        
        XWPFTable table = doc.getTableArray(2);
        
        int charge_rows = charge_materials.size();
        
        int aa = 1;
        int k = 5;
        int init = (charge_rows > 13) ? 13 : charge_rows;
        int i;
        
        for(i = 0; i < init; i++)
        {   
            XWPFParagraph paragraph = doc.createParagraph();
            XWPFRun paragraphRun = paragraph.createRun();
            
            paragraphRun.setFontSize(7);
            paragraphRun.setFontFamily("Arial");
            paragraphRun.setText(String.valueOf(aa) + ".");
            
            XWPFTableRow tableRow = table.getRow(k);
            tableRow.getCell(0).setParagraph(paragraph);
            paragraph.removeRun(0);
           
            aa++;
            
            for (int j = 0; j < 6; j++)
            {
                XWPFParagraph par = doc.createParagraph();
                XWPFRun parRun = par.createRun();
                
                parRun.setFontSize(7);
                parRun.setFontFamily("Arial");
                parRun.setText(charge_materials.get(i)[(j < 5) ? j + 1 : 5].toString());
                
                XWPFTableRow tRow = table.getRow(k);
                tRow.getCell(j + 1).setParagraph(par);
                par.removeRun(0);
            }
        
            k++;
        }
        
        for(i = 13; i < charge_rows; i++)
        {
             try
             {
                 XWPFTableRow lastRow = table.getRows().get(table.getNumberOfRows() - 1);
                 CTRow ctrow = CTRow.Factory.parse(lastRow.getCtRow().newInputStream());
                 XWPFTableRow newRow = new XWPFTableRow(ctrow, table);
                 
                XWPFParagraph paragraph = doc.createParagraph();
                XWPFRun paragraphRun = paragraph.createRun();

                paragraphRun.setFontSize(7);
                paragraphRun.setFontFamily("Arial");
                paragraphRun.setText(String.valueOf(aa) + ".");
                
                aa++;

                newRow.getCell(0).setParagraph(paragraph);
                paragraph.removeRun(0);

                 for (int j = 0; j < 6; j++)
                 {
                     XWPFParagraph par = doc.createParagraph();
                     XWPFRun parRun = par.createRun();

                     parRun.setFontSize(7);
                     parRun.setFontFamily("Arial");
                     parRun.setText(charge_materials.get(i)[j + 1].toString());

                     newRow.getCell(j + 1).setParagraph(par);
                     par.removeRun(0);
                 }

                 table.addRow(newRow);
             }
             catch(Exception e)
             {
                 System.out.println("Error adding more rows");
             }
         }
    
    }
    
    private void writeAll()
    {
        writeField(From, 0);
        writeField(To, 1);
        writeDate();
        writeTable();
    }
    
    public void createDocx()
    {
        writeAll();
        
        try (FileOutputStream out = new FileOutputStream(Filename + ".docx"))
        {
            doc.write(out);
        }
        catch(IOException e)
        {
            System.out.println("Error writing 108 docx.");
        }
    }
}
