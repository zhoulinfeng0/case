package com.example;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

public class ReplacePPTAttachment2 {
    public static void main(String[] args) throws Exception {
        // 1. 加载 PPT 文件
        String pptFilePath = "example.pptx";
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(pptFilePath));

        // 2. 遍历所有幻灯片，找到嵌入的文件
        List<XSLFSlide> slides = ppt.getSlides();
        for (XSLFSlide slide : slides) {
            // 遍历幻灯片中的所有相关部件
            for (POIXMLDocumentPart part : slide.getRelations()) {
                PackagePart packagePart = part.getPackagePart();
                String contentType = packagePart.getContentType();

                // 检查嵌入对象（通常是 OLE 对象）
                if (contentType.contains("vnd.openxmlformats-officedocument.oleObject")) {
                    System.out.println("找到 OLE 对象：" + packagePart.getPartName());

                    // 替换嵌入文件内容
                    FileInputStream newAttachment = new FileInputStream(new File("newAttachment.xls")); // 新文件
                    try (FileOutputStream outputStream = (FileOutputStream) packagePart.getOutputStream()) {
                        byte[] buffer = new byte[1024];
                        int bytesRead;
                        while ((bytesRead = newAttachment.read(buffer)) != -1) {
                            outputStream.write(buffer, 0, bytesRead);
                        }
                    }
                    newAttachment.close();
                    System.out.println("附件替换成功：" + packagePart.getPartName());
                }
            }
        }

        // 3. 保存修改后的 PPT
        try (FileOutputStream out = new FileOutputStream("modified.pptx")) {
            ppt.write(out);
        }
        ppt.close();

        System.out.println("PPT 文件修改完成！");
    }
}

