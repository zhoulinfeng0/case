package com.example;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.xslf.usermodel.XMLSlideShow;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

public class ReEmbedPptObject {
    public static void main(String[] args) throws Exception {
        String pptFilePath = "src/main/resources/table.pptx";
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(pptFilePath));

        // 获取所有嵌入对象
        List<PackagePart> embeddedParts = ppt.getAllEmbeddedParts();
        if (embeddedParts.isEmpty()) {
            System.out.println("No embedded parts found.");
            return;
        }

        // 获取第一个嵌入对象
        PackagePart packagePart = embeddedParts.get(0);

        System.out.println("Part Name: " + packagePart.getPartName());
        System.out.println("Content Type: " + packagePart.getContentType());

        // 获取 OPCPackage
        OPCPackage opcPackage = ppt.getPackage();

        // 删除旧的嵌入对象（检查是否存在）
        if (opcPackage.getPart(packagePart.getPartName()) != null) {
            opcPackage.deletePart(packagePart.getPartName());
            System.out.println("Deleted part: " + packagePart.getPartName());
        } else {
            System.out.println("Part not found for deletion: " + packagePart.getPartName());
        }

        // 创建新的嵌入对象
        String newFilePath = "src/main/resources/test2.xlsx"; // 替换为新的文件路径
        File newFile = new File(newFilePath);
        String contentType = "application/vnd.ms-excel"; // 根据文件类型设置

        PackagePart newPart = opcPackage.createPart(
                PackagingURIHelper.createPartName("/ppt/embeddings/newObject1.bin"), contentType
        );

        // 写入新的文件内容
        try (OutputStream os = newPart.getOutputStream();
             FileInputStream fis = new FileInputStream(newFile)) {
            byte[] buffer = new byte[1024];
            int bytesRead;
            while ((bytesRead = fis.read(buffer)) != -1) {
                os.write(buffer, 0, bytesRead);
            }
        }

        // 保存修改后的文件
        FileOutputStream out = new FileOutputStream("modified.pptx");
        ppt.write(out);
        ppt.close();
        out.close();

        System.out.println("File modified successfully.");
    }
}


