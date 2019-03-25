package data;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;


public class ProtocolDocCreator
{
    public static void main(String[] args) throws Exception
    {          
        FileInputStream fis;
        
        XWPFDocument doc = new XWPFDocument (fis = new FileInputStream("src/rsc/protocol.docx"));
        
        //imerominia
        String date = "15/1/19";
        XWPFTable table0 = doc.getTableArray(0);

        XWPFParagraph paragraph0 = doc.createParagraph();
        XWPFRun paragraphOneRun0 = paragraph0.createRun();
        paragraphOneRun0.setFontSize(12);
        paragraphOneRun0.setFontFamily("Arial");
        paragraphOneRun0.setText(date);
        XWPFTableRow tableRow0 = table0.getRow(0);
        tableRow0.getCell(0).setParagraph(paragraph0);
        paragraph0.removeRun(0);
        
        //diatagi
        String order = "kurwakurwakurwakurwakurwa";
        XWPFTable table1 = doc.getTableArray(1);
        
        XWPFParagraph paragraph1 = doc.createParagraph();
        XWPFRun paragraphOneRun1 = paragraph1.createRun();
        paragraphOneRun1.setFontSize(12);
        paragraphOneRun1.setFontFamily("Arial");
        paragraphOneRun1.setText(order);
        XWPFTableRow tableRow1 = table1.getRow(0);
        tableRow1.getCell(0).setParagraph(paragraph1);
        paragraph1.removeRun(0);
        
        //proedros
        String proed = "proedors";
        XWPFTable table2 = doc.getTableArray(2);
        
        XWPFParagraph paragraph2 = doc.createParagraph();
        XWPFRun paragraphOneRun2 = paragraph2.createRun();
        paragraphOneRun2.setFontSize(12);
        paragraphOneRun2.setFontFamily("Arial");
        paragraphOneRun2.setText(proed);
        XWPFTableRow tableRow2 = table2.getRow(0);
        tableRow2.getCell(0).setParagraph(paragraph2);
        paragraph2.removeRun(0);
        
        //melos 1
        String melos1 = "loche";
        XWPFTable table3 = doc.getTableArray(3);
        
        XWPFParagraph paragraph3 = doc.createParagraph();
        XWPFRun paragraphOneRun3 = paragraph3.createRun();
        paragraphOneRun3.setFontSize(12);
        paragraphOneRun3.setFontFamily("Arial");
        paragraphOneRun3.setText(melos1);
        XWPFTableRow tableRow3 = table3.getRow(0);
        tableRow3.getCell(0).setParagraph(paragraph3);
        paragraph3.removeRun(0);
        
        //melos 2
        String melos2 = "treloche";
        XWPFTable table4 = doc.getTableArray(4);
        
        XWPFParagraph paragraph4 = doc.createParagraph();
        XWPFRun paragraphOneRun4 = paragraph4.createRun();
        paragraphOneRun4.setFontSize(12);
        paragraphOneRun4.setFontFamily("Arial");
        paragraphOneRun4.setText(melos2);
        XWPFTableRow tableRow4 = table4.getRow(0);
        tableRow4.getCell(0).setParagraph(paragraph4);
        paragraph4.removeRun(0);
        
        
        //ilika
        String ilika[][] = new String [][] {  {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"},
                                              {"kappa","keepo","kurwa","bifkiri","amakerasoun"}
                                           };
        
        int rows = 22;
        int aa = 1;
        int k = 2;
        
        XWPFTable table5 = doc.getTableArray(6);
        
        
        if(rows > table5.getNumberOfRows()-2)
        {
           int test = rows-(table5.getNumberOfRows()-2);
           for(int i = 0; i < test; i++)
           {
                //insert new last row, which is a copy previous last row:
               XWPFTableRow lastRow = table5.getRows().get(table5.getNumberOfRows() - 1);
               CTRow ctrow = CTRow.Factory.parse(lastRow.getCtRow().newInputStream());
               XWPFTableRow newRow = new XWPFTableRow(ctrow, table5);

               for (XWPFTableCell cell : newRow.getTableCells()) {
                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                 for (XWPFRun run : paragraph.getRuns()) {
                  run.setText("", 0);
                 }
                }
               }

               table5.addRow(newRow);
           }
        }
        XWPFTable table6 = doc.getTableArray(6);
        System.out.println(table6.getNumberOfRows());
        for(int i = 0; i < rows; i++)
        {
            XWPFParagraph paragraph5 = doc.createParagraph();
            XWPFRun paragraphOneRun5 = paragraph5.createRun();
            paragraphOneRun5.setFontSize(12);
            paragraphOneRun5.setFontFamily("Arial");
            paragraphOneRun5.setText(String.valueOf(aa));
            XWPFTableRow tableRow5 = table6.getRow(k);
            tableRow5.getCell(0).setParagraph(paragraph5);
            paragraph5.removeRun(0);
            
            aa++;
            for(int j = 0; j < 5; j++)
            {
                XWPFParagraph par5 = doc.createParagraph();
                XWPFRun paragraphOne5 = par5.createRun();
                paragraphOne5.setFontSize(12);
                paragraphOne5.setFontFamily("Arial");
                paragraphOne5.setText(ilika[i][j]);
                XWPFTableRow tableRow6 = table6.getRow(k);
                tableRow6.getCell(j+1).setParagraph(par5);
                par5.removeRun(0);
            }
            k++;
        }
     
        try (FileOutputStream out = new FileOutputStream("F:/NetBeansProjects/DataWarehouse/out.docx"))
        {
            doc.write(out);
        }
        catch(IOException e)
        {
            System.out.println("error " +e);
        }
        
    }
}