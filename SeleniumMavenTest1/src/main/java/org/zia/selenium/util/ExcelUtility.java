package org.zia.selenium.util;

import jxl.Workbook;
import jxl.Sheet;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;

public class ExcelUtility {

	public String testcase;
	public WritableSheet writableSheet;
	public WritableWorkbook writableWorkbook;

	@BeforeTest
	public void queryParameterization() throws BiffException, IOException, RowsExceededException, WriteException {

		FileInputStream testInput = new FileInputStream(
				"C:\\Users\\Tariq Ahsan\\Desktop\\ISB Java JEE Training\\Selenium\\TestData\\InputDataFile.xls");

		// Now get Workbook
		// Workbook workbook = Workbook.getWorkbook(testFile);
		Workbook workbook = Workbook.getWorkbook(testInput);

		// Now get Workbook Sheet
		// Sheet sheet = workbook.getSheet("query_data");
		Sheet sheets = workbook.getSheet("query_data");
		// Sheet sheet = workbook.getSheet(0);. Workbook.createWorkbook(new
		// File(outputfileName));
		// WritableSheet ws = wwb.createSheet("topicName", 0);

		int numOfRows = sheets.getRows();

		// Read rows and columns and save it in Two Dimentional String Array
		String inputdata[][] = new String[sheets.getRows()][sheets.getColumns()];

		System.out.println("Number of rows in the testData.xls file is -> " + numOfRows);

		// For writing the data in a Excel spreadsheet we will FileOutputStream
		FileOutputStream testOutput = new FileOutputStream(
				"C:\\Users\\Tariq Ahsan\\Desktop\\ISB Java JEE Training\\Selenium\\TestData\\InputDataFile1.xls");
		System.out.println("Creating a file ...");

		// Create writable workbook
		writableWorkbook = Workbook.createWorkbook(testOutput);

		// Create writable sheet
		writableSheet = writableWorkbook.createSheet("query_data", 0);

		// Write all data in new sheet using for loop
		for (int i = 0; i < sheets.getRows(); i++) {
			for (int j = 0; j < sheets.getColumns(); j++) {
				inputdata[i][j] = sheets.getCell(j, i).getContents();
				Label label1 = new Label(j, i, inputdata[i][j]);
				Label label2 = new Label(4, 0, "Results");
				writableSheet.addCell(label1);
				writableSheet.addCell(label2);
			}
		}

	}

	@AfterTest
	public void writeExcels() throws WriteException {
		try {
			writableWorkbook.write();
		} catch (IOException ioEx) {
			ioEx.printStackTrace();
		}
	}
}
