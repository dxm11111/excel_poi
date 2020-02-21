package com.atguigu.excelpoi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;

public class ExcelReadTest {
    //HSSFWorkbook Excel03(.xls) 读文件
    @Test
    public void testRead03() throws IOException {

        InputStream is =  new FileInputStream("D:/190805/excel_poi/商品表-03.xls");

        Workbook wb = new HSSFWorkbook(is);
        //Sheet sheet = wb.getSheet("类别");
        //根据索引来读取到sheet
        Sheet sheet = wb.getSheetAt(0);

        //获取到第0行
        Row row = sheet.getRow(0);
        //获取到第0列
        Cell cell = row.getCell(0);

        String cellValue = cell.getStringCellValue();
        System.out.println(cellValue);

    }


    //XSSFWorkbook Excel07(.xls) 读文件
    @Test
    public void testRead07() throws IOException {

        InputStream is =  new FileInputStream("D:/190805/excel_poi/商品表-07.xlsx");

        Workbook wb = new XSSFWorkbook(is);
        //Sheet sheet = wb.getSheet("类别");
        //根据索引来读取到sheet
        Sheet sheet = wb.getSheetAt(0);

        //获取到第0行
        Row row = sheet.getRow(0);
        //获取到第0列
        Cell cell = row.getCell(0);

        String cellValue = cell.getStringCellValue();
        System.out.println(cellValue);

    }

    //HSSFWorkbook Excel03(.xls) 读文件
    @Test
    public void testReadCellType() throws IOException {

        InputStream is =  new FileInputStream("D:/190805/excel_poi/会员消费商品明细表.xls");

        Workbook wb = new HSSFWorkbook(is);
        //Sheet sheet = wb.getSheet("类别");
        //根据索引来读取到sheet
        Sheet sheet = wb.getSheetAt(0);

        //获取到第0行
        //Row row = sheet.getRow(0);
        //获取到第0列
        //Cell cell = row.getCell(0);

        //读取标题行
        Row rowTitle = sheet.getRow(0);
        //判断标题行是否为空
        if(rowTitle != null){
            //获取单元格数量
            int cellCount = rowTitle.getPhysicalNumberOfCells();

            for (int cellNum = 0; cellNum < cellCount; cellNum++){
                Cell cell = rowTitle.getCell(cellNum);
                //判断单元格是否为空
                if(cell != null){
                    int cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + "-" + cellType + " ");
                }
            }
            System.out.println();
        }

        //读取数据行
        int rowCount = sheet.getPhysicalNumberOfRows();
        for(int rowNum = 1; rowNum < rowCount; rowNum++){
            //获取数据行对象
            Row rowData = sheet.getRow(rowNum);
            if(rowData != null){
                int cellCount = rowData.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCount;cellNum++){
                    //获取单元格数量
                    Cell cell = rowData.getCell(cellNum);
                    if(cell != null){

                        //获取单元格数据类型
                         int cellType = cell.getCellType();
                         String cellValue = "";
                        System.out.print("【" + (rowNum + 1) + "-" + (cellNum + 1) +"】");

                        HSSFFormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook)wb);

                        switch (cellType) {
                            case HSSFCell .CELL_TYPE_STRING:

                                System.out.print("【STRING】");
                                cellValue = cell.getStringCellValue();

                                break;
                            case HSSFCell .CELL_TYPE_NUMERIC:
                                System.out.print("【NUMERIC】");

                                if (DateUtil.isCellDateFormatted(cell)) {
                                    Date date = cell.getDateCellValue();
                                    System.out.print(new DateTime(date).toString("yyyy-MM-dd"));
                                } else {
                                    //将当前单元格转化成字符串类型
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    System.out.print(cell.getStringCellValue());
                                }
                                break;
                            case HSSFCell .CELL_TYPE_BOOLEAN:
                                System.out.print("【BOOLEAN】");
                                System.out.print(cell.getBooleanCellValue());
                                break;
                            case HSSFCell .CELL_TYPE_FORMULA:
                                System.out.print("【FORMULA】");

                                CellValue eveluate = formulaEvaluator.evaluate(cell);
                                cellValue = eveluate.formatAsString();
                                System.out.print(cellValue);
                                break;
                            case HSSFCell.CELL_TYPE_BLANK:
                                System.out.print("【BLANK】");
                                System.out.print("");
                                break;
                            case HSSFCell.CELL_TYPE_ERROR:
                                System.out.print("【ERROR】");
                                System.out.println("【数据类型错误】");
                                break;
                            default:
                                System.out.println();
                         }

                         System.out.println(cellValue);
                    }
                }
            }

        }
        is.close();
    }


    //读取公式
    @Test
    public void testReadFormula() throws IOException {

        InputStream is =  new FileInputStream("D:/190805/excel_poi/计算公式.xls");

        Workbook wb = new HSSFWorkbook(is);
        //Sheet sheet = wb.getSheet("类别");
        //根据索引来读取到sheet
        Sheet sheet = wb.getSheetAt(0);

        //获取到第0行
        Row row = sheet.getRow(4);
        //获取到第0列
        Cell cell = row.getCell(0);

        String cellValueFormula = cell.getCellFormula();
        HSSFFormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook)wb);
        CellValue eveluate = formulaEvaluator.evaluate(cell);
        String cellValue = eveluate.formatAsString();
        System.out.println(cellValueFormula);
        System.out.println(cellValue);

    }

}
