package com.zhang.easy;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.util.Date;

@Data
public class DemoData {
//    @ExcelProperty("字符串标题")
    @ExcelProperty("姓名")
    private String string;
//    @ExcelProperty("日期标题")
    @ExcelProperty("出生日期")
    private Date date;
//    @ExcelProperty("数字标题")
    @ExcelProperty("体重")
    private Double doubleData;
    /**
     * 忽略这个字段
     */
    @ExcelIgnore
    private String ignore;
}