package com.kuang;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.FileInputStream;

/**
 * @author wuyz1
 */
public class ExcelReadTest {

    String PATH = "D:\\code\\excel\\";

    @Test
    public void testRead03() throws Exception {
        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH+"poi-excel狂神观众统计表03.xls");
        //创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        //1.得到表    可以根据下标取sheet，也可以通过名字取sheet
        Sheet sheet = workbook.getSheetAt(0);
        //2.得到行
        Row row = sheet.getRow(0);
        //3.得到列
        Cell cell = row.getCell(0);
        String cellValue = cell.getStringCellValue();
        System.out.println(cellValue);


    }
}
