package com.example.util;

import java.io.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class CompressionUtil {
    
    public static class CompressionResult {
        private long originalSize;
        private long compressedSize;
        private long timeSpent;
        private double compressionRatio;

        public CompressionResult(long originalSize, long compressedSize, long timeSpent) {
            this.originalSize = originalSize;
            this.compressedSize = compressedSize;
            this.timeSpent = timeSpent;
            this.compressionRatio = (1 - ((double) compressedSize / originalSize)) * 100;
        }

        @Override
        public String toString() {
            return String.format(
                "压缩结果统计:\n" +
                "原始大小: %d 字节\n" +
                "压缩后大小: %d 字节\n" +
                "压缩率: %.2f%%\n" +
                "耗时: %d 毫秒",
                originalSize, compressedSize, compressionRatio, timeSpent
            );
        }
    }

    /**
     * 压缩文件或目录
     * @param sourcePath 源文件或目录路径
     * @param targetPath 目标zip文件路径
     * @return 压缩结果统计
     */
    public static CompressionResult compress(String sourcePath, String targetPath) throws IOException {
        File sourceFile = new File(sourcePath);
        long startTime = System.currentTimeMillis();
        long originalSize = calculateSize(sourceFile);
        
        try (FileOutputStream fos = new FileOutputStream(targetPath);
             ZipOutputStream zos = new ZipOutputStream(fos)) {
            
            compressFile(sourceFile, sourceFile.getName(), zos);
        }
        
        long endTime = System.currentTimeMillis();
        long compressedSize = new File(targetPath).length();
        
        return new CompressionResult(originalSize, compressedSize, endTime - startTime);
    }

    /**
     * 解压文件
     * @param zipPath zip文件路径
     * @param targetDir 解压目标目录
     * @return 解压结果统计
     */
    public static CompressionResult decompress(String zipPath, String targetDir) throws IOException {
        long startTime = System.currentTimeMillis();
        File zipFile = new File(zipPath);
        long compressedSize = zipFile.length();
        
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(zipFile))) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                File file = new File(targetDir, entry.getName());
                
                if (entry.isDirectory()) {
                    file.mkdirs();
                    continue;
                }
                
                file.getParentFile().mkdirs();
                try (FileOutputStream fos = new FileOutputStream(file)) {
                    byte[] buffer = new byte[1024];
                    int len;
                    while ((len = zis.read(buffer)) > 0) {
                        fos.write(buffer, 0, len);
                    }
                }
            }
        }
        
        long endTime = System.currentTimeMillis();
        long originalSize = calculateSize(new File(targetDir));
        
        return new CompressionResult(originalSize, compressedSize, endTime - startTime);
    }

    private static void compressFile(File file, String fileName, ZipOutputStream zos) throws IOException {
        if (file.isDirectory()) {
            File[] files = file.listFiles();
            if (files != null) {
                for (File child : files) {
                    compressFile(child, fileName + "/" + child.getName(), zos);
                }
            }
            return;
        }

        try (FileInputStream fis = new FileInputStream(file)) {
            zos.putNextEntry(new ZipEntry(fileName));
            byte[] buffer = new byte[1024];
            int len;
            while ((len = fis.read(buffer)) > 0) {
                zos.write(buffer, 0, len);
            }
            zos.closeEntry();
        }
    }

    private static long calculateSize(File file) {
        if (file.isFile()) {
            return file.length();
        }
        
        long size = 0;
        File[] files = file.listFiles();
        if (files != null) {
            for (File child : files) {
                size += calculateSize(child);
            }
        }
        return size;
    }

    // 使用示例
    public static void main(String[] args) {
        try {
            // 压缩文件
            CompressionResult compressResult = compress("./source", "./compressed.zip");
            System.out.println("压缩完成：");
            System.out.println(compressResult);

            // 解压文件
            CompressionResult decompressResult = decompress("./compressed.zip", "./extracted");
            System.out.println("\n解压完成：");
            System.out.println(decompressResult);
            
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
} 