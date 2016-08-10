package com.sai.excelr;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelDemo {
	static int x=0;
	
	public static String domainName(String str){
		String[] st = new String[2]; 
		st=str.split("@");
		
		
		return "www."+st[1];
	}
	public static void main(String[] args) 
    {
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		Map<String,String> map = new TreeMap<String, String>();
		String companyName=" ";
		
		  try
	        {
	            FileInputStream file1 = new FileInputStream(new File("f:\\excelapp\\DiscoverOrgCompanyDetail6316.xlsx"));
	            //Create Workbook instance holding reference to .xlsx file
	            XSSFWorkbook workbook3 = new XSSFWorkbook(file1);
	 
	            //Get first/desired sheet from the workbook
	            XSSFSheet sheet3 = workbook3.getSheetAt(0);
	 
	            //Iterate through each rows one by one
	            Iterator<Row> rowIterator1 = sheet3.iterator();
	            while (rowIterator1.hasNext()) 
	            {
	                Row row = rowIterator1.next();
	                //For each row, iterate through all the columns
	              //  Iterator<Cell> cellIterator = row.cellIterator();
	                Cell c =  row.getCell(2);
	                Cell c1 = row.getCell(1);
	                
	                
	                
	               
	                
	                
	                if(c == null )
	                	continue;
	                
	              //  System.out.println(c.getStringCellValue()+" ------------------- "+c1.getStringCellValue());
	                
	                map.put(c.getStringCellValue(), c1.getStringCellValue());
	                
	               
	            }
	            System.out.println("--------------");
	            file1.close();
	        } 
	        catch (Exception e) 
	        {
	            e.printStackTrace();
	        }
		  
		  	System.out.println(map.containsKey("www.directv.com"));
		  	System.out.println(map.get("www.directv.com"));
		  
		
       try
        {
            FileInputStream file = new FileInputStream(new File("f:\\excelapp\\DiscoverOrgPersonReport.xlsx"));
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
 
            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
 
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) 
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                Cell c =  row.getCell(6);
                Cell c1 = row.getCell(0);
                
                
                if(c == null )
                	continue;
                
               // System.out.print(x++ +"\t");
               // System.out.print(c1.getStringCellValue()+"----------");
              //  System.out.println(c.getStringCellValue()+"\t");
               // System.out.println(domainName(c.getStringCellValue()));
               
              String check = map.get(domainName(c.getStringCellValue()));
            //  System.out.println(check);
              if(check != null){
            	  System.out.println(check);
            	  data.put(Integer.toString(x++), new Object[] { c1.getStringCellValue(),c.getStringCellValue(),domainName(c.getStringCellValue()),check });
            	  
              }
              else{
            	  data.put(Integer.toString(x++), new Object[] { c1.getStringCellValue(),c.getStringCellValue(),domainName(c.getStringCellValue()),"No Data Found" });
              }
             
                
                
              //  data.put(Integer.toString(x++), new Object[] { c1.getStringCellValue(),c.getStringCellValue(),domainName(c.getStringCellValue()),companyName });
                
                
               
            }
            file.close();
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
        XSSFWorkbook workbook1 = new XSSFWorkbook(); 
        XSSFSheet writeSheet = workbook1.createSheet("Data");
        
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset)
        {
            Row row = writeSheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
               Cell cell = row.createCell(cellnum++);
               if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("f:\\excelapp\\test_data.xlsx"));
            workbook1.write(out);
            out.close();
            System.out.println("success");
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
        
    }
}


