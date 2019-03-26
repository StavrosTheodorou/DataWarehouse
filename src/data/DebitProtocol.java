package data;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class DebitProtocol 
{
    private final String President;
    private final String Order;
    private final String Member1;
    private final String Member2;
    private String Date;
    private final String Filename;
    private final ArrayList<Object[]> charge_materials;
    private XWPFDocument doc = null;

    public DebitProtocol(String President, String Order, String Member1, String Member2, String Date, String Filename, ArrayList<Object[]> charge_materials) 
    {
        this.President = President;
        this.Order = Order;
        this.Member1 = Member1;
        this.Member2 = Member2;
        this.Date = Date;
        this.Filename = Filename;
        this.charge_materials = charge_materials;
        
        convertDate();
        
        try
        {
            doc = new XWPFDocument(new FileInputStream("src/rsc/protocol.docx"));
        }
        catch(Exception e)
        {
            System.out.println("Error creating protocol docx.");
        }
    }
    
    private void convertDate()
    {
        SimpleDateFormat format = new SimpleDateFormat("dd/MM/yyyy");
        
        
        try
        {
            Date date = format.parse(Date);
            String Month = "", Day = "";
            
            switch(date.getMonth())
            {
                case 0: Month = "Ιανουαρίου"; break;
                case 1: Month = "Φεβρουαρίου"; break;
                case 2: Month = "Μαρτίου"; break;
                case 3: Month = "Απριλίου"; break;
                case 4: Month = "Μαΐου"; break;
                case 5: Month = "Ιουνίου"; break;
                case 6: Month = "Ιουλίου"; break;
                case 7: Month = "Αυγούστου"; break;
                case 8: Month = "Σεπτεμβρίου"; break;
                case 9: Month = "Οκτωβρίου"; break;
                case 10: Month = "Νοεμβρίου"; break;
                case 11: Month = "Δεκεμβίου"; break;
            }
            
            switch(date.getDay())
            {
                case 0: Day = "Κυριακή"; break;
                case 1: Day = "Δευτέρα"; break;
                case 2: Day = "Τρίτη"; break;
                case 3: Day = "Τετάρτη"; break;
                case 4: Day = "Πέμπτη"; break;
                case 5: Day = "Παρασκευή"; break;
                case 6: Day = "Σάββατο"; break;
            }
            
            Date = date.getDate() + " " + Month + " " + (date.getYear() + 1900) + " ημέρα " + Day;
        }
        catch(Exception e)
        {
            System.out.println("Error parsing date");
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
    
    private void writeTable()
    {
        /*
        //Debugging
        String ilika[][] = new String [][] {  {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"},
                                              {"test","test","test","test","test"}
                                           };
        */
   
        XWPFTable table = doc.getTableArray(6);
        
        int charge_rows = charge_materials.size();
        int table_rows = table.getNumberOfRows() - 2;
        
        int aa = 1;
        int k = 2;
        
        if(charge_rows > table_rows)
        {
           int size = charge_rows - table_rows;
           
           for(int i = 0; i < size; i++) table.createRow();
        }
        
        System.out.println("Table rows: " + (table.getRows().size() - 2) );
        
        for(int i = 0; i < charge_rows; i++)
        {
            XWPFParagraph paragraph = doc.createParagraph();
            XWPFRun paragraphRun = paragraph.createRun();
            
            paragraphRun.setFontSize(12);
            paragraphRun.setFontFamily("Arial");
            paragraphRun.setText(String.valueOf(aa));
            
            XWPFTableRow tableRow = table.getRow(k);
            tableRow.getCell(0).setParagraph(paragraph);
            paragraph.removeRun(0);
            
            aa++;
            
            for(int j = 0; j < 5; j++)
            {
                XWPFParagraph par= doc.createParagraph();
                XWPFRun parRun = par.createRun();
                
                parRun.setFontSize(12);
                parRun.setFontFamily("Arial");
                parRun.setText(charge_materials.get(i)[convertIndex(j)].toString());
                
                XWPFTableRow tRow = table.getRow(k);
                tRow.getCell(j + 1).setParagraph(par);
                par.removeRun(0);
            }
            
            k++;
        }
    }
    
    private int convertIndex(int j)
    {
        switch(j)
        {
            default: return -1;
            
            case 0: return 2;
            case 1: return 3;
            case 2: return 4;
            case 3: return 5;
            case 4: return 7;
        }
    }
    
    private void writeAll()
    {
        writeField(Date, 0);
        writeField(Order, 1);
        writeField(President, 2);
        writeField(Member1, 3);
        writeField(Member2, 4);
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
            System.out.println("Error writing protocol docx.");
        }
    }
}
