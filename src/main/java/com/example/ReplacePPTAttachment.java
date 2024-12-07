package com.example;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.openxml4j.opc.PackagePart;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

public class ReplacePPTAttachment {
    public static void main(String[] args) throws Exception {
        // 1. 加载PPT文件
        String pptFilePath = "src/main/resources/template1.pptx";
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(pptFilePath));

        // 2. 遍历所有幻灯片
        List<XSLFSlide> slides = ppt.getSlides();
        for (XSLFSlide slide : slides) {
            // 遍历每个幻灯片的嵌入对象
            for (PackagePart part : slide.getPackagePart().getPackage().getParts()) {
                if (part.getPartName().getName().endsWith(".xlsx") || part.getPartName().getName().endsWith(".pdf")) {
                    // 找到要替换的附件（例如 Excel 或 PDF）
                    System.out.println("找到嵌入文件：" + part.getPartName().getName());

                    // 3. 替换附件内容
                    FileInputStream newAttachment = new FileInputStream(new File("/Users/mac/IdeaProjects/g-ppt-util/src/main/resources/1.xlsx")); // 替换的新文件
                    OutputStream partOutputStream = part.getOutputStream();
                    byte[] buffer = new byte[1024];
                    int bytesRead;
                    while ((bytesRead = newAttachment.read(buffer)) != -1) {
                        partOutputStream.write(buffer, 0, bytesRead);
                    }
                    partOutputStream.close();
                    newAttachment.close();
                }
            }
        }
        //代码写的不错哦 ，哈哈
        //这样就舒服了 ，哈哈～

        // 4. 保存修改后的PPT
        FileOutputStream out = new FileOutputStream("modified4.pptx");
        ppt.write(out);
        out.close();
        ppt.close();

        System.out.println("附件替换完成！");



     }
}
