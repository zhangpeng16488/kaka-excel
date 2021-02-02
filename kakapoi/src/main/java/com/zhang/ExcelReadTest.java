package com.zhang;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * @auther kaka
 * @create 2021-02-02 23:18
 */
public class ExcelReadTest {

    String path = "E:\\develop\\kaka-excel\\kakapoi\\";

    @Test
    public void testRead03() throws Exception {

        FileInputStream inputStream = new FileInputStream(path + "kaka.xls");

        //创建工作簿
        Workbook workbook = new HSSFWorkbook(inputStream);
        //创建工作表
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        String value = cell.getStringCellValue();
        System.out.println(value);
    }
    @Test
    public void testRead07() throws Exception {

        FileInputStream inputStream = new FileInputStream(path + "kaka.xlsx");

        //创建工作簿
        Workbook workbook = new XSSFWorkbook(inputStream);
        //创建工作表
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        String value = cell.getStringCellValue();
        System.out.println(value);
    }
}
