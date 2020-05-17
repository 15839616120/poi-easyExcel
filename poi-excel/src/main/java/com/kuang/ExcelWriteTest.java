package com.kuang;

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

/**
 * @author wuyz1
 */
public class ExcelWriteTest {

    String PATH = "D:\\code\\poi-easyExcel\\poi-excel";

    @Test
    public void testWrite03() throws Exception {
        //创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        //创建一个工作表
        Sheet sheet = workbook.createSheet("狂神观众统计表");
        //创建行
        Row row1 = sheet.createRow(0);
        //创建行内的单元格(1,1)
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增观众");
        //创建行内的单元格(1,2)
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);

        Row row2 = sheet.createRow(1);
        //(2,1)
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计日期");
        //(2,2)
        Cell cell22 = row2.createCell(1);
        cell22.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        //生产一张表
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "狂神观众统计表03.xls");
        workbook.write(fileOutputStream);

        //关闭流
        System.out.println("03生成完毕");

    }

    @Test
    public void testWrite07() throws Exception {
        //创建一个工作簿
        Workbook workbook = new XSSFWorkbook();
        //创建一个工作表
        Sheet sheet = workbook.createSheet("狂神观众统计表");
        //创建行
        Row row1 = sheet.createRow(0);
        //创建行内的单元格(1,1)
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增观众");
        //创建行内的单元格(1,2)
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(666);

        Row row2 = sheet.createRow(1);
        //(2,1)
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计日期");
        //(2,2)
        Cell cell22 = row2.createCell(1);
        cell22.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        //生产一张表
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "狂神观众统计表07.xlsx");
        workbook.write(fileOutputStream);

        //关闭流
        System.out.println("07生成完毕");

    }
    public static void main(String[] args) {

        String isInvestor="111";

        System.out.println( "0".equals(isInvestor)?"":"updateSetType");
    }
    @Test
    public void testWrite03BigData() throws Exception {
        //开始时间
        long begin = System.currentTimeMillis();
        //创建一个工作薄
        Workbook workbook = new SXSSFWorkbook();
        //创建一张表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0;rowNum < 65536; rowNum++){
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0;cellNum < 10; cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        //输出
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWrite03BigData.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        //结束时间
        long end = System.currentTimeMillis();
        System.out.println(begin-end);

    }
}
