package com.example;

import java.awt.*;

public class FontsTest {

    public static void main(String[] args) {
        String[] fonts = GraphicsEnvironment
                .getLocalGraphicsEnvironment()
                .getAvailableFontFamilyNames();

        System.out.println("系统中的中文字体：");
        for (String font : fonts) {
            // 筛选可能的中文字体
            if (font.contains("GB") ||
                    font.contains("Ming") ||
                    font.contains("Song") ||
                    font.contains("Hei") ||
                    font.contains("Kai") ||
                    font.contains("Fang") ||
                    font.contains("华文") ||
                    font.contains("宋") ||
                    font.contains("黑") ||
                    font.contains("门") ||
                    font.contains("楷")) {
//                System.out.println(font);
            }else {
                System.out.println(font);

            }
        }
    }
}
