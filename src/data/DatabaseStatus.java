package data;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class DatabaseStatus 
{
    private final String Filename;
    private final ArrayList<Object[]> materials;
    private final String holder;
    private XWPFDocument doc = null;

    public DatabaseStatus(String Filename, ArrayList<Object[]> materials, String holder) 
    {
        this.Filename = Filename;
        this.materials = materials;
        this.holder = holder;
        
        try
        {            
            doc = new XWPFDocument(getClass().getClassLoader().getResourceAsStream("rsc/status.docx"));
        }
        catch(Exception e)
        {
            System.out.println("Error creating protocol docx.");
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
        paragraphRun.setBold(true);
        
        XWPFTableRow tableRow = table.getRow(0);
        tableRow.getCell(0).setParagraph(paragraph);
        paragraph.removeRun(0);
    }
    
    private void writeTable()
    {
        XWPFTable table = doc.getTableArray(2);
             
        for (int i = 0; i < materials.size(); i++)
        {            
            table.createRow();
            
            for (int j = 0; j < materials.get(0).length - 1; j++)
            {                
                XWPFParagraph paragraph = doc.createParagraph();
                XWPFRun paragraphRun = paragraph.createRun();

                paragraphRun.setFontSize(12);
                paragraphRun.setFontFamily("Arial");
                paragraphRun.setText(materials.get(i)[j].toString());

                XWPFTableRow tableRow = table.getRow(i + 1);
                tableRow.getCell(j).setParagraph(paragraph);
                paragraph.removeRun(0);
            }
        }
        
    }
    
    private void removeEmptyLines()
    {
        int last_element_index = doc.getBodyElements().size() - 1;
        
        do
        {
            IBodyElement be = doc.getBodyElements().get(last_element_index);
            
            if (be.getElementType() != BodyElementType.PARAGRAPH) break;
            
            XWPFParagraph p = (XWPFParagraph) be;
            
            if (!p.getText().equals("")) break;
            
            doc.removeBodyElement(last_element_index--);
        }
        while(true);
    }   
    
    private void writeAll()
    {
        writeField("Ημερομηνία: " + DateTimeFormatter.ofPattern("dd/MM/yyyy").format(LocalDateTime.now()), 0);
        writeField("Κάτοχος: " + holder, 1);
        writeTable();
        removeEmptyLines();
    }
         
    public void createDocx()
    {
        writeAll();
        
        new File("./ΚΑΤΑΣΤΑΤΙΚΑ").mkdirs();
        
        try (FileOutputStream out = new FileOutputStream("./ΚΑΤΑΣΤΑΤΙΚΑ/" + Filename + ".docx"))
        {
            doc.write(out);
        }
        catch(IOException e)
        {
            System.out.println("Error writing status docx.");
        }
    }
}
