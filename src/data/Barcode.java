package data;

import com.itextpdf.text.pdf.Barcode39;
import java.awt.Color;
import java.awt.Font;
import java.awt.FontMetrics;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.util.Random;

public class Barcode 
{
    private static Random r = new Random();
    
    private String digits;
    private BufferedImage bimg;
    
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
        bimg = drawTextOnImage("* " + digits + " *", toBufferedImage(bc.createAwtImage(Color.BLACK, Color.WHITE)));
    }
    
    private BufferedImage drawTextOnImage(String text, BufferedImage image) 
    {
        BufferedImage bi = new BufferedImage(image.getWidth(), image.getHeight() + 20, BufferedImage.TRANSLUCENT);
        Graphics2D g2d = (Graphics2D) bi.createGraphics();
        g2d.addRenderingHints(new RenderingHints(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY));
        g2d.addRenderingHints(new RenderingHints(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON));
        g2d.addRenderingHints(new RenderingHints(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON));

        g2d.drawImage(image, 0, 0, null);

        g2d.setColor(Color.BLACK);
        g2d.setFont(new Font("Calibri", Font.BOLD, 20));
        FontMetrics fm = g2d.getFontMetrics();
        int textWidth = fm.stringWidth(text);

        //center text at bottom of image in the new space
        g2d.drawString(text, (bi.getWidth() / 2) - textWidth / 2, bi.getHeight());

        g2d.dispose();
        return bi;
    }
    
    private BufferedImage toBufferedImage(Image img)
    {
        // Create a buffered image with transparency
        BufferedImage bimage = new BufferedImage(img.getWidth(null), img.getHeight(null), BufferedImage.TYPE_INT_ARGB);

        // Draw the image on to the buffered image
        Graphics2D bGr = bimage.createGraphics();
        bGr.drawImage(img, 0, 0, null);
        bGr.dispose();

        // Return the buffered image
        return bimage;
    }
    
    public Image getImg() { return bimg; }
    
    @Override
    public String toString()
    {
        return digits;
    }
}
