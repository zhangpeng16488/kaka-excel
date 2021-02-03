package com.zhang;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Date;

/**
 * @auther kaka
 * @create 2021-02-02 23:18
 */
public class ExcelReadTest {

    //    String path = "E:\\develop\\kaka-excel\\kakapoi\\";
    String path = "D:\\Idea-Workspace\\excel-test\\kakapoi\\";

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

    @Test
    public void testCellType() throws Exception{

        FileInputStream inputStream = new FileInputStream(path + "明细表.xls");
        Workbook workbook = new HSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        Row rowTitle = sheet.getRow(0);
        if(rowTitle != null){
            int cells = rowTitle.getPhysicalNumberOfCells();
            for (int i = 0; i < cells; i++) {
                Cell cell = rowTitle.getCell(i);
                if(cell != null){
//                    int cellType = cell.getCellType();
//                    System.out.print(cellType + " | ");
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + " | ");
                }
            }
            System.out.println();
        }

        int rows = sheet.getPhysicalNumberOfRows();
        for (int i = 1; i < rows; i++) {
            Row row = sheet.getRow(i);
            if(row != null){
                int number = row.getPhysicalNumberOfCells();
                for (int j = 0; j < number; j++) {
//                    System.out.print("[" + (i + 1) + "-" + (j + 1) + "]");

                    Cell cell = row.getCell(j);
                    if(cell != null){
                        int cellType = cell.getCellType();
                        String cellValue = "";
                        switch (cellType){
                            case Cell.CELL_TYPE_STRING:
//                                System.out.println("[String]");
                                cellValue = cell.getStringCellValue();
                                break;
                            case Cell.CELL_TYPE_BOOLEAN:
//                                System.out.println("[String]");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case Cell.CELL_TYPE_BLANK:
//                                System.out.println("[BLANK]");
                                break;
                            case Cell.CELL_TYPE_NUMERIC:
//                                System.out.println("[NUMERIC]");
                                if(HSSFDateUtil.isCellDateFormatted(cell)){
//                                    System.out.println("日期");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd");
                                }else {
//                                    System.out.println("转换为字符串输出");
                                    cell.setCellType(Cell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case Cell.CELL_TYPE_ERROR:
//                                System.out.println("数据类型错误");
                                break;
                        }
                        System.out.print(cellValue + "|");
                    }
                }
            }
            System.out.println();
        }
        inputStream.close();
    }

    @Test
    public void testFormula()throws Exception{
        FileInputStream inputStream = new FileInputStream(path + "公式.xls");
        Workbook workbook = new HSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(4);
        Cell cell = row.getCell(0);
        FormulaEvaluator evaluator = new HSSFFormulaEvaluator((HSSFWorkbook)workbook);

        int cellType = cell.getCellType();
        switch (cellType){
            case Cell.CELL_TYPE_FORMULA:
                String formula = cell.getCellFormula();
                System.out.println(formula);
                CellValue evaluate = ((HSSFFormulaEvaluator) evaluator).evaluate(cell);
                String format = evaluate.formatAsString();
                System.out.println(format);
                break;
        }
    }
}
