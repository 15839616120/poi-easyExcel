package com.kuang;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.joda.time.DateTime;

/**
 * @author wuyz1
 */
public class ExcelWriteTest {


    public void testWrite03() {
        //创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        //创建一个工作表
        Sheet sheet = workbook.createSheet("狂神观众统计表");
        //创建行
        Row row = sheet.createRow(0);
        //创建行内的单元格(1,1)
        Cell cell11 = row.createCell(0);
        cell11.setCellValue("今日新增观众");
        //创建行内的单元格(1,2)
        Cell cell12 = row.createCell(1);
        cell12.setCellValue(666);

        Row row2 = sheet.createRow(1);
        //(2,1)
        Cell cell21 = row.createCell(0);
        cell21.setCellValue("统计日期");
        cell21.setCellValue(new DateTime().toString());

    }

    public static void main(String[] args) {

        String isInvestor="111";

        System.out.println( "0".equals(isInvestor)?"":"updateSetType");
    }
}
