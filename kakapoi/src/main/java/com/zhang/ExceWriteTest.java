package com.zhang;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

public class ExceWriteTest {

    String path = "E:\\develop\\kaka-excel\\kakapoi\\";

    @Test
    public void testWrite03() throws Exception {

        //创建工作簿
        Workbook workbook = new HSSFWorkbook();
        //创建工作表
        Sheet sheet = workbook.createSheet("狂神观众统计表");
        Row row1 = sheet.createRow(0);
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增观众");
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);

        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成excel，excel03结尾是xls
        FileOutputStream outputStream = new FileOutputStream(path + "kaka.xls");
        workbook.write(outputStream);
        outputStream.close();
        System.out.println("kaka.xls生成完毕");
    }
    @Test
    public void testWrite07() throws Exception {

        //创建工作簿
        Workbook workbook = new XSSFWorkbook();
        //创建工作表
        Sheet sheet = workbook.createSheet("狂神观众统计表");
        Row row1 = sheet.createRow(0);
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增观众");
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);

        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        Cell cell22 = row2.createCell(1);
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        //生成excel，excel07结尾是xlsx
        FileOutputStream outputStream = new FileOutputStream(path + "kaka.xlsx");
        workbook.write(outputStream);
        outputStream.close();
        System.out.println("kaka.xlsx生成完毕");
    }
    //最多65536行
    @Test
    public void testWrite03BigData() throws Exception {

        long start = System.currentTimeMillis();
        //创建工作簿
        Workbook workbook = new HSSFWorkbook();
        //创建工作表
        Sheet sheet = workbook.createSheet();

        for (int i = 0; i < 65536; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(j);
            }
        }

        //生成excel，excel03结尾是xls
        FileOutputStream outputStream = new FileOutputStream(path + "kaka03BigData.xls");
        workbook.write(outputStream);
        outputStream.close();
        System.out.println("kaka03BigData.xls生成完毕");
        long end = System.currentTimeMillis();
        System.out.println((end - start) / 1000);
    }
    //耗时较长，可以写入1048576
    @Test
    public void testWrite07BigData() throws Exception {

        long start = System.currentTimeMillis();
        //创建工作簿
        Workbook workbook = new XSSFWorkbook();
        //创建工作表
        Sheet sheet = workbook.createSheet();

        for (int i = 0; i < 65536; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(j);
            }
        }

        //生成excel，excel03结尾是xls
        FileOutputStream outputStream = new FileOutputStream(path + "kaka07BigData.xlsx");
        workbook.write(outputStream);
        outputStream.close();
        System.out.println("kaka07BigData.xlsx生成完毕");
        long end = System.currentTimeMillis();
        System.out.println((end - start) / 1000);
    }
    @Test
    public void testWrite07BigDataSuper() throws Exception {

        long start = System.currentTimeMillis();
        //创建工作簿
        Workbook workbook = new SXSSFWorkbook();
        //创建工作表
        Sheet sheet = workbook.createSheet();

        for (int i = 0; i < 100000; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(j);
            }
        }

        //生成excel，excel03结尾是xls
        FileOutputStream outputStream = new FileOutputStream(path + "kaka07BigDataSuper.xlsx");
        workbook.write(outputStream);
        outputStream.close();
        System.out.println("kaka07BigDataSuper.xlsx生成完毕");
        //清除临时文件
        ((SXSSFWorkbook)workbook).dispose();
        long end = System.currentTimeMillis();
        System.out.println((end - start) / 1000);
    }
}