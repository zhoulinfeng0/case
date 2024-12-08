package com.example;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import javax.imageio.ImageIO;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import com.itextpdf.text.Document;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.pdf.PdfWriter;

public class PPTToPDFConverter {
    
    /**
     * 将PPT转换为PDF
     * @param inputPath PPT文件路径
     * @param outputPath 输出PDF路径
     * @param scale 分辨率缩放比例，建议值2-4
     * @throws Exception 转换过程中可能出现的异常
     */
    public static void convertPPTtoPDF(String inputPath, String outputPath, double scale) throws Exception {
        // 加载PPT文件
        File pptFile = new File(inputPath);
        try (XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(pptFile))) {
            // 创建PDF文档
            Document pdfDocument = new Document(PageSize.A4.rotate());
            PdfWriter.getInstance(pdfDocument, new FileOutputStream(outputPath));
            pdfDocument.open();

            // 获取PPT尺寸
            Dimension pgsize = ppt.getPageSize();
            int width = (int) (pgsize.width * scale);
            int height = (int) (pgsize.height * scale);

            // 临时文件
            File tempImg = new File("temp-slide.png");
            try {
                // 处理每一页
                for (XSLFSlide slide : ppt.getSlides()) {
                    convertSlideToPDF(slide, width, height, scale, tempImg, pdfDocument);
                }
            } finally {
                // 清理临时文件
                if (tempImg.exists()) {
                    tempImg.delete();
                }
                pdfDocument.close();
            }
        }
    }

    private static void convertSlideToPDF(XSLFSlide slide, int width, int height, double scale, 
                                        File tempImg, Document pdfDocument) throws Exception {
        BufferedImage img = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D graphics = img.createGraphics();
        try {
            // 设置渲染质量
            configureGraphics(graphics, width, height);
            
            // 设置中文字体
            setChineseFont(graphics);

            // 应用缩放
            graphics.scale(scale, scale);

            // 渲染幻灯片
            slide.draw(graphics);

            // 保存为临时图片
            ImageIO.write(img, "png", tempImg);

            // 添加到PDF
            addImageToPDF(tempImg, pdfDocument);
        } finally {
            graphics.dispose();
        }
    }

    private static void configureGraphics(Graphics2D graphics, int width, int height) {
        // 设置渲染质量
        graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
        graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
        graphics.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
        graphics.setRenderingHint(RenderingHints.KEY_STROKE_CONTROL, RenderingHints.VALUE_STROKE_PURE);

        // 设置背景色
        graphics.setColor(Color.white);
        graphics.fill(new Rectangle2D.Float(0, 0, width, height));
    }

    private static void setChineseFont(Graphics2D graphics) {
        try {
            String[] fontFamilies = GraphicsEnvironment
                .getLocalGraphicsEnvironment()
                .getAvailableFontFamilyNames();
            
            String[] preferredFonts = {
                    "Menlo","宋体", "SimSun", "微软雅黑", "Microsoft YaHei", "黑体", "SimHei"};
            Font selectedFont = null;
            
            for (String preferredFont : preferredFonts) {
                for (String family : fontFamilies) {
                    if (family.equals(preferredFont)) {
                        selectedFont = new Font(family, Font.PLAIN, 12);
                        break;
                    }
                }
                if (selectedFont != null) break;
            }
            
            if (selectedFont == null) {
                selectedFont = new Font(Font.SANS_SERIF, Font.PLAIN, 12);
            }
            
            graphics.setFont(selectedFont);
        } catch (Exception e) {
            e.printStackTrace();
            graphics.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 12));
        }
    }

    private static void addImageToPDF(File tempImg, Document pdfDocument) throws Exception {
        Image pdfImage = Image.getInstance(tempImg.getAbsolutePath());
        float documentWidth = pdfDocument.getPageSize().getWidth() - 40;
        float documentHeight = pdfDocument.getPageSize().getHeight() - 40;
        pdfImage.scaleToFit(documentWidth, documentHeight);
        pdfImage.setAlignment(Image.MIDDLE);
        pdfDocument.add(pdfImage);
    }

    // 使用示例
    public static void main(String[] args) {
//        test1();
//        test2();
//        test3();
//        test4();
        test5();
    }

    public static void test1(){
        try {
            String inputPath = "src/main/resources/极简素雅灰色通用PPT模板.pptx";
            String outputPath = "极简素雅灰色通用PPT模板.pdf";
            convertPPTtoPDF(inputPath, outputPath, 2.0);
            System.out.println("转换完成！");
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("转换失败：" + e.getMessage());
        }
    }

    public static void test2(){
        try {
            String inputPath = "src/main/resources/淡雅清新唯美花朵PPT模板.pptx";
            String outputPath = "淡雅清新唯美花朵PPT模板.pdf";
            convertPPTtoPDF(inputPath, outputPath, 2.0);
            System.out.println("转换完成！");
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("转换失败：" + e.getMessage());
        }
    }

    public static void test3(){
        try {
            String inputPath = "src/main/resources/小清新文艺范LOMO风PPT模板.pptx";
            String outputPath = "小清新文艺范LOMO风PPT模板.pdf";
            convertPPTtoPDF(inputPath, outputPath, 2.0);
            System.out.println("转换完成！");
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("转换失败：" + e.getMessage());
        }
    }

    public static void test4(){
        try {
            String inputPath = "src/main/resources/箭头简约工作总结计划PPT模板.pptx";
            String outputPath = System.currentTimeMillis()+ ".pdf";
            convertPPTtoPDF(inputPath, outputPath, 2.0);
            System.out.println("转换完成！");
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("转换失败：" + e.getMessage());
        }
    }

    public static void test5(){
        try {
            String inputPath = "src/main/resources/2020-02-18_1582013313_5e4b9b819f414.pptx";
            String outputPath = System.currentTimeMillis()+ ".pdf";
            convertPPTtoPDF(inputPath, outputPath, 2.0);
            System.out.println("转换完成！");
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("转换失败：" + e.getMessage());
        }
    }
}
