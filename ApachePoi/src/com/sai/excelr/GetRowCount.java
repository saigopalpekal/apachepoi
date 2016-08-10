package com.sai.excelr;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class GetRowCount {
	
	static int RowCount=0;
	public static void main(String[] args) throws InvalidFormatException, EncryptedDocumentException, IOException {
		
		FileInputStream fis = new FileInputStream("f:\\excelapp\\DiscoverOrgPersonReport.xlsx");
		
		Workbook wb = WorkbookFactory.create(fis);
		org.apache.poi.ss.usermodel.Sheet s = wb.getSheet("DiscoverOrg_Person_121611_20160");
		
		RowCount = s.getLastRowNum();
		System.out.println(RowCount);
	}

}
