package org.zia.selenium.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelDataConfig {
	
	// Accessible within the class
	HSSFWorkbook wb;
	HSSFSheet sheet1;
	
	public ExcelDataConfig(String excelPath) {
		try {
			
			// Only accessible within the method
			File src = new File(excelPath);
			FileInputStream fis = new FileInputStream(src);
		    wb = new HSSFWorkbook(fis);
			
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	
	}
	
	public String getData(int sheetNumber, int row, int column) {
		
		sheet1 = wb.getSheetAt(sheetNumber);
		String data = sheet1.getRow(row).getCell(column).getStringCellValue();
		return data;
	}
	
	public HSSFSheet getSheetNumber(int i) {
		return sheet1 = wb.getSheetAt(i);
	}
	
	public int getRowCount(int sheetIndex) {
		int row = wb.getSheetAt(sheetIndex).getLastRowNum();
		return row + 1;
		
	}

}
