//package com.example;import org.apache.poi.xslf.usermodel.*;
//import org.apache.poi.sl.usermodel.*;
//
//import java.io.*;
//import javax.imageio.ImageIO;
//import java.awt.image.BufferedImage;
//
//public class PPTExcelReplacement {
//
//    public static void main(String[] args) throws Exception {
//        String pptxFilePath = "src/main/resources/table.pptx";
//        String imagePath = "src/main/resources/test2.xlsx"; // 图像路径
//        String outputPptxFilePath = "output_pptx_file.pptx";
//
//        // 读取 PPTX 文件
//        FileInputStream pptxFile = new FileInputStream(pptxFilePath);
//        XMLSlideShow ppt = new XMLSlideShow(pptxFile);
//
//        // 读取图像文件
//        FileInputStream imageFile = new FileInputStream(imagePath);
//        BufferedImage bufferedImage = ImageIO.read(imageFile);
//
//        // 处理 PPT 文件中的每一页
//        for (XSLFSlide slide : ppt.getSlides()) {
//            for (XSLFShape shape : slide.getShapes()) {
//                // 检查是否为嵌入的 OLE 对象
//                if (shape instanceof XSLFObjectShape) {
//                    XSLFObjectShape oleObjectShape = (XSLFObjectShape) shape;
//                    System.out.println(oleObjectShape.getFullName());
//
//                    slide.removeShape(oleObjectShape);  // 删除原有的 OLE 对象
//                    insertImageAsReplacement(slide, bufferedImage);  // 插入新的图像替代
//                }
//            }
//        }
//
//        // 保存修改后的 PPT 文件
//        FileOutputStream out = new FileOutputStream(outputPptxFilePath);
//        ppt.write(out);
//        out.close();
//        pptxFile.close();
//
//        System.out.println("OLE object replaced with image successfully!");
//    }
//
//    // 插入新的图像替代 OLE 对象
//    private static void insertImageAsReplacement(XSLFSlide slide, BufferedImage bufferedImage) throws IOException {
//        // 将图像转换为字节流
//        ByteArrayOutputStream imageStream = new ByteArrayOutputStream();
//        ImageIO.write(bufferedImage, "PNG", imageStream);
//        byte[] imageBytes = imageStream.toByteArray();
//
//        // 插入图像
//        XSLFPictureShape pictureShape = slide.cr(pictureData);
//
//        // 设置图像的位置和大小
//        pictureShape.setAnchor(new java.awt.Rectangle(100, 100, 400, 300)); // 调整位置和大小
//    }
//}
//
