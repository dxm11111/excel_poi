package com.atguigu.excelpoi;

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
import java.io.IOException;
import java.util.Date;
//写文件 Excel03(.xls)
public class ExcelWriteTest {

    @Test
    public void testWrite03() throws IOException {

        Workbook wb = new HSSFWorkbook();
        Sheet sheet1 = wb.createSheet("new sheet");
        Sheet sheet2 = wb.createSheet("second sheet");

        Row row = sheet1.createRow(0);
        Cell cell1 = row.createCell(0);
        Cell cell2 = row.createCell(1);
        Cell cell3 = row.createCell(2);

        cell1.setCellValue("hello excel");
        cell2.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));
        cell3.setCellValue("12345678911");
         FileOutputStream out = new FileOutputStream("D:/190805/excel_poi/test-write03.xls");
         wb.write(out);

         out.close();

         System.out.println("文件生成成功");
    }

    //写文件 Excel07(.xlsx)
    @Test
    public void testWrite07() throws IOException {

        Workbook wb = new XSSFWorkbook();
        Sheet sheet1 = wb.createSheet("new sheet");
        Sheet sheet2 = wb.createSheet("second sheet");

        Row row = sheet1.createRow(0);
        Cell cell1 = row.createCell(0);
        Cell cell2 = row.createCell(1);
        Cell cell3 = row.createCell(2);

        cell1.setCellValue("hello excel");
        cell2.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));
        cell3.setCellValue("12345678911");
        FileOutputStream out = new FileOutputStream("D:/190805/excel_poi/test-write07.xlsx");
        wb.write(out);

        out.close();

        System.out.println("文件生成成功");
    }
    //HSSFWorkbook Excel03写文件
    //65536行耗时：2.092秒 16824k
    //如果大于65536行就会抛出异常，来防止OOM     java.lang.IllegalArgumentException: Invalid row number (65536) outside allowable range (0..65535)
    @Test
    public void testWrite03BigData() throws IOException {
        //记录开始时间
        long begin = System.currentTimeMillis();

        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet();

        //循环创建65536行记录
        for(int rowNum = 0;rowNum < 65536;rowNum++ ){
            Row row = sheet.createRow(rowNum);
            //一共有10个单元格
            for (int cellNum = 0; cellNum < 10; cellNum++){
                Cell cell = row.createCell(cellNum);
                //要写的内容
                cell.setCellValue(rowNum + "-" + cellNum);
            }
        }
        //写文件
        FileOutputStream out = new FileOutputStream("D:/190805/excel_poi/test-write03-bigdata.xls");
        wb.write(out);
        out.close();
        System.out.println("文件写完毕");

        //记录结束时间
        long end = System.currentTimeMillis();

        //耗时
        System.out.println("耗时："+ (double)(end-begin)/1000);
    }


    //XSSFWorkbook Excel07写文件
    //65536行耗时：15.738秒 3688k
    //大于65536行  可以写出来   大约1000000 会发生OOM
    @Test
    public void testWrite07BigData() throws IOException {
        //记录开始时间
        long begin = System.currentTimeMillis();

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet();

        //循环创建65536行记录
        for(int rowNum = 0;rowNum < 65536;rowNum++ ){
            Row row = sheet.createRow(rowNum);
            //一共有10个单元格
            for (int cellNum = 0; cellNum < 10; cellNum++){
                Cell cell = row.createCell(cellNum);
                //要写的内容
                cell.setCellValue(rowNum + "-" + cellNum);
            }
        }
        //写文件
        FileOutputStream out = new FileOutputStream("D:/190805/excel_poi/test-write07-bigdata.xlsx");
        wb.write(out);
        out.close();
        System.out.println("文件写完毕");

        //记录结束时间
        long end = System.currentTimeMillis();

        //耗时
        System.out.println("耗时："+ (double)(end-begin)/1000);
    }


    //SXSSFWorkbook Excel07快速写文件    既节省时间又节省内存   几乎不会发生OOM
    @Test
    public void testWrite07BigDataFast() throws IOException {
        //记录开始时间
        long begin = System.currentTimeMillis();

        Workbook wb = new SXSSFWorkbook();
        Sheet sheet = wb.createSheet();

        //循环创建65536行记录
        for(int rowNum = 0;rowNum < 65536;rowNum++ ){
            Row row = sheet.createRow(rowNum);
            //一共有10个单元格
            for (int cellNum = 0; cellNum < 10; cellNum++){
                Cell cell = row.createCell(cellNum);
                //要写的内容
                cell.setCellValue(rowNum + "-" + cellNum);
            }
        }
        //写文件
        FileOutputStream out = new FileOutputStream("D:/190805/excel_poi/test-write07-bigdata-fast.xlsx");
        wb.write(out);
        out.close();
        System.out.println("文件写完毕");

        //清除临时文件
        ((SXSSFWorkbook)wb).dispose();

        //记录结束时间
        long end = System.currentTimeMillis();

        //耗时
        System.out.println("耗时："+ (double)(end-begin)/1000);
    }
}
