package com.example;


 import java.io.File;
 import java.io.FileInputStream;
 import java.io.FileOutputStream;
 import java.io.IOException;
 import java.util.List;
 import java.util.Map;
 import org.apache.poi.xslf.usermodel.XMLSlideShow;
 import org.apache.poi.xslf.usermodel.XSLFTable;
 import org.apache.poi.xslf.usermodel.XSLFTableRow;
 import org.apache.poi.xslf.usermodel.XSLFTextShape;

 import java.util.Map;

public class PptTemplateParser {

    // 加载模板
    public XMLSlideShow loadTemplate(String templatePath) {
        // 从路径加载 PPT 模板
        try (FileInputStream fis = new FileInputStream(templatePath)) {
            return new XMLSlideShow(fis);
        } catch (IOException e) {
            throw new RuntimeException("模板加载失败：" + e.getMessage(), e);
        }
    }

    // 填充数据并生成新文件
    public void fillTemplate(XMLSlideShow ppt, Map<String, Object> data, String outputPath) {
        // 遍历幻灯片并替换占位符
        ppt.getSlides().forEach(slide -> {
            slide.getShapes().forEach(shape -> {
                if (shape instanceof XSLFTextShape) {
                    XSLFTextShape textShape = (XSLFTextShape) shape;
                    String text = textShape.getText();
                    for (Map.Entry<String, Object> entry : data.entrySet()) {
                        text = text.replace("${" + entry.getKey() + "}", entry.getValue().toString());
                    }
                    textShape.setText(text);
                }


                if (shape instanceof XSLFTable) {
                    XSLFTable textShape = (XSLFTable) shape;
                    List<XSLFTableRow> rows = textShape.getRows();

                }
            });
        });

        // 保存到指定路径
        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
            ppt.write(fos);
        } catch (IOException e) {
            throw new RuntimeException("文件保存失败：" + e.getMessage(), e);
        }
    }
}

