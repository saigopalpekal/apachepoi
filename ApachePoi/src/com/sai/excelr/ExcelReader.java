package com.sai.excelr;

import java.io.*;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class ExcelReader 
{
static void readXlsx(File inputFile) 
{
try 
{
        // Get the workbook instance for XLSX file
        XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(inputFile));

        // Get first sheet from the workbook
        XSSFSheet sheet = wb.getSheetAt(0);

        Row row;
        Cell cell;

        // Iterate through each rows from first sheet
        Iterator<Row> rowIterator = sheet.iterator();

        while (rowIterator.hasNext()) 
        {
                row = rowIterator.next();

                // For each row, iterate through each columns
                Iterator<Cell> cellIterator = row.cellIterator();
                
                while (cellIterator.hasNext()) 
                {
                cell = cellIterator.next();

                switch (cell.getCellType()) 
                {

                case Cell.CELL_TYPE_BOOLEAN:
                        System.out.println(cell.getBooleanCellValue());
                        break;

                case Cell.CELL_TYPE_NUMERIC:
                        System.out.println(cell.getNumericCellValue());
                        break;

                case Cell.CELL_TYPE_STRING:
                        System.out.println(cell.getStringCellValue());
                        break;

                case Cell.CELL_TYPE_BLANK:
                        System.out.println(" ");
                        break;

                default:
                        System.out.println(cell);

                }
                }
        }
}
catch (Exception e) 
{
        System.err.println("Exception :" + e.getMessage());
}
}

static void readXls(File inputFile) 
{
try 
{
        // Get the workbook instance for XLS file
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(inputFile));
        // Get first sheet from the workbook
        HSSFSheet sheet = workbook.getSheetAt(0);
        Cell cell;
        Row row;

        // Iterate through each rows from first sheet
        Iterator<Row> rowIterator = sheet.iterator();
        
        while (rowIterator.hasNext()) 
        {
                row = rowIterator.next();

                // For each row, iterate through each columns
                Iterator<Cell> cellIterator = row.cellIterator();
                
                while (cellIterator.hasNext()) 
                {
                cell = cellIterator.next();

                switch (cell.getCellType()) 
                {

                case Cell.CELL_TYPE_BOOLEAN:
                        System.out.println(cell.getBooleanCellValue());
                        break;

                case Cell.CELL_TYPE_NUMERIC:
                        System.out.println(cell.getNumericCellValue());
                        break;

                case Cell.CELL_TYPE_STRING:
                        System.out.println(cell.getStringCellValue());
                        break;

                case Cell.CELL_TYPE_BLANK:
                        System.out.println(" ");
                        break;

                default:
                        System.out.println(cell);
                }
                }
        }

} 

catch (FileNotFoundException e) 
{
        System.err.println("Exception" + e.getMessage());
}
catch (IOException e) 
{
        System.err.println("Exception" + e.getMessage());
}
}

public static void main(String[] args) 
{
       // File inputFile = new File("C:\input.xls");
        File inputFile2 = new File("f:\\excelapp\\DiscoverOrgPersonReport.xlsx");
        //readXls(inputFile);
        readXlsx(inputFile2);
}
}