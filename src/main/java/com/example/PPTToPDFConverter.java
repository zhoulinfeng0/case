package com.example;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import javax.imageio.ImageIO;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import com.itextpdf.text.Document;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.pdf.PdfWriter;

public class PPTToPDFConverter {
    public static void main(String[] args) throws Exception {
        // Load PPTX file
        File pptFile = new File("src/main/resources/template.pptx");
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(pptFile));

        // Create PDF document
        Document pdfDocument = new Document();
        PdfWriter.getInstance(pdfDocument, new FileOutputStream("output.pdf"));
        pdfDocument.open();

        // Convert each slide to an image
        Dimension pageSize = ppt.getPageSize();
        for (XSLFSlide slide : ppt.getSlides()) {
            BufferedImage img = new BufferedImage(pageSize.width, pageSize.height, BufferedImage.TYPE_INT_RGB);
            Graphics2D graphics = img.createGraphics();

            // Render slide content to image
            slide.draw(graphics);
            graphics.dispose();

            // Save image temporarily
            File tempImg = new File("temp-slide.jpg");
            ImageIO.write(img, "jpg", tempImg);

            // Add image to PDF
            Image pdfImage = Image.getInstance(tempImg.getAbsolutePath());
            pdfImage.scaleToFit(PageSize.A4.getWidth(), PageSize.A4.getHeight());
            pdfDocument.add(pdfImage);

            tempImg.delete(); // Clean up
        }

        pdfDocument.close();
        ppt.close();
    }
}
