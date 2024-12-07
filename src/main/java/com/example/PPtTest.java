package com.example;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

public class PPtTest {
    public static void main(String[] args) {
        Map<String, Object> data = new HashMap<>();
        data.put("title", "这是标题");
        data.put("tableData", Arrays.asList(
                Arrays.asList("列1", "列2", "列3"),
                Arrays.asList("数据1", "数据2", "数据3")
        ));


        PptTemplateParser parser = new PptTemplateParser();
        XMLSlideShow ppt = parser.loadTemplate("/Users/mac/IdeaProjects/g-ppt-util/src/main/resources/template.pptx");
        parser.fillTemplate(ppt, data, "outp2ut.pptx");
    }

}
