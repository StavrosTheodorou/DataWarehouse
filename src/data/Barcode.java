package data;

import com.itextpdf.text.pdf.Barcode39;
import java.awt.Color;
import java.awt.Image;
import java.util.Random;

public class Barcode 
{
    private static Random r = new Random();
    
    private String digits;
    private Image img;
    
    public Barcode()
    {
        generateDigits();
        createImage();
    }
    
    public Barcode(String digits)
    {
        this.digits = digits;
        createImage();
    }
    
    private void generateDigits()
    {
        digits = "";
        
        for (int i = 0; i < 8; i++) digits += r.nextInt(9);
    }
    
    private void createImage()
    {       
        Barcode39 bc = new Barcode39();
        bc.setCode(digits);
        bc.setBarHeight(50);
        img = bc.createAwtImage(Color.BLACK, Color.WHITE);
    }
    
    public Image getImg() { return img; }
    
    @Override
    public String toString()
    {
        return digits;
    }
}
