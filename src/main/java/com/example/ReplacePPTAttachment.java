package com.example;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.openxml4j.opc.TargetMode;

import java.io.*;
import java.util.List;

public class ReplacePPTAttachment {
    public static void main(String[] args) {
        FileInputStream pptFileStream = null;
        FileOutputStream outStream = null;
        XMLSlideShow ppt = null;

        try {
            // 1. 加载PPT文件
            String pptFilePath = "src/main/resources/table.pptx";
            pptFileStream = new FileInputStream(pptFilePath);
            ppt = new XMLSlideShow(pptFileStream);

            boolean foundAnyExcel = false;
            // 2. 遍历所有幻灯片
            List<XSLFSlide> slides = ppt.getSlides();
            for (XSLFSlide slide : slides) {
                // 获取幻灯片中的所有关系
                PackageRelationshipCollection rels = slide.getPackagePart().getRelationships();
                
                for (PackageRelationship rel : rels) {
                    // 只处理内部嵌入的对象
                    if (rel.getTargetMode() != TargetMode.INTERNAL) {
                        continue;
                    }

                    try {
                        PackagePart part = slide.getPackagePart().getRelatedPart(rel);
                        String contentType = part.getContentType().toLowerCase();
                        String partName = part.getPartName().getName().toLowerCase();
                        
                        System.out.println("检查对象: " + partName);
                        System.out.println("Content Type: " + contentType);
                        
                        // 检查是否是嵌入的Excel文件
                        if (isEmbeddedExcel(contentType, partName, rel.getRelationshipType())) {
                            System.out.println("找到嵌入的Excel文件：" + part.getPartName().getName());
                            System.out.println("关系类型: " + rel.getRelationshipType());

                            // 3. 替换附件内容
                            replaceAttachment(part, "src/main/resources/test2.xlsx");
                            foundAnyExcel = true;
                        }
                    } catch (Exception e) {
                        System.err.println("处理关系 " + rel.getId() + " 时发生错误: " + e.getMessage());
                        // 继续处理下一个关系
                        continue;
                    }
                }
            }

            if (!foundAnyExcel) {
                System.out.println("警告：未找到任何可替换的Excel文件！");
                return;
            }

            // 4. 保存修改后的PPT
            outStream = new FileOutputStream("modified_compatible.pptx");
            ppt.write(outStream);
            System.out.println("附件替换完成！");

        } catch (Exception e) {
            System.err.println("处理PPT文件时发生错误: " + e.getMessage());
            e.printStackTrace();
        } finally {
            // 5. 关闭所有资源
            closeQuietly(outStream);
            closeQuietly(pptFileStream);
            closeQuietly(ppt);
        }
    }

    private static boolean isEmbeddedExcel(String contentType, String partName, String relationType) {
        // Office中常见的Excel嵌入对象的内容类型
        String[] excelContentTypes = {
            "application/vnd.openxmlformats-officedocument.spreadsheetml",
            "application/vnd.ms-excel",
            "application/x-ole-storage",
            "application/vnd.openxmlformats-officedocument.oleObject"
        };

        // 检查关系类型
        boolean isOleObject = relationType != null && 
            (relationType.contains("oleObject") || relationType.contains("package"));

        // 检查文件扩展名
        boolean hasExcelExtension = partName.endsWith(".xlsx") || 
                                  partName.endsWith(".xls") || 
                                  partName.contains("excel");

        // 检查内容类型
        boolean hasExcelContentType = false;
        for (String type : excelContentTypes) {
            if (contentType.contains(type.toLowerCase())) {
                hasExcelContentType = true;
                break;
            }
        }

        return (hasExcelContentType || hasExcelExtension || isOleObject);
    }

    private static void replaceAttachment(PackagePart part, String newFilePath) throws IOException {
        File newFile = new File(newFilePath);
        if (!newFile.exists()) {
            throw new FileNotFoundException("替换用的Excel文件不存在: " + newFilePath);
        }

        try (FileInputStream newAttachment = new FileInputStream(newFile);
             OutputStream partOutputStream = part.getOutputStream()) {
            
            byte[] buffer = new byte[8192];
            int bytesRead;
            while ((bytesRead = newAttachment.read(buffer)) != -1) {
                partOutputStream.write(buffer, 0, bytesRead);
            }
            partOutputStream.flush();
        }
    }

    private static void closeQuietly(Closeable resource) {
        if (resource != null) {
            try {
                resource.close();
            } catch (IOException e) {
                System.err.println("关闭资源时发生错误: " + e.getMessage());
            }
        }
    }
}
