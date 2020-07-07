package com.prograd.delete;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Delete {

	public static void main(String[] args) throws IOException {
	File myFile = new File("D:\\Round_3_Result.xlsx"); 
	
	FileInputStream fis = new FileInputStream(myFile); // Finds the workbook instance for XLSX file 
	
	XSSFWorkbook myWorkBook = new XSSFWorkbook (fis); // Return first sheet from the XLSX workbook 
	
	XSSFSheet mySheet = myWorkBook.getSheetAt(0); // Get iterator to all the rows in current sheet 
	
	Row row = mySheet.getRow(0);
	row.removeCell(row.getCell(0));
	
	FileOutputStream os = new FileOutputStream(myFile);
    myWorkBook.write(os);
    System.out.println("Delete Done");
	}
}
