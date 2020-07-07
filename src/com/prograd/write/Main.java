package com.prograd.write;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		File myFile = new File("D:\\Round_3_Result.xlsx"); 
		
		FileInputStream fis = new FileInputStream(myFile); // Finds the workbook instance for XLSX file 
		
		XSSFWorkbook myWorkBook = new XSSFWorkbook (fis); // Return first sheet from the XLSX workbook 
		
		XSSFSheet mySheet = myWorkBook.getSheetAt(0); // Get iterator to all the rows in current sheet 
		
	        Object[][] Rows = {
	              
	        		{"Datatype", "Type", "Size(in bytes)"},
	                {"int", "Primitive", 2},
	                {"float", "Primitive", 4},
	                {"double", "Primitive", 8},
	                {"char", "Primitive", 1},
	                {"String", "Non-Primitive", "No fixed size"}
	        };

	        int rowNum = mySheet.getLastRowNum(); //1

	        for (Object[] s_row : Rows) { //{"", "Type", "Size(in bytes)"},
	        	
	            Row row = mySheet.createRow(rowNum++);//1++
	            int colNum = 0;
	            
	            for (Object cell_value : s_row) {
	                Cell cell = row.createCell(colNum++);//1
	                if (cell_value instanceof String) {
	                    cell.setCellValue((String) cell_value);
	                } else if (cell_value instanceof Integer) {
	                    cell.setCellValue((Integer) cell_value);
	                }else if (cell_value instanceof Boolean) {
	                    cell.setCellValue((Boolean) cell_value);
	                }else if (cell_value instanceof Date) {
	                    cell.setCellValue((Date) cell_value);
	                }
	            }
	        }

	        FileOutputStream os = new FileOutputStream(myFile);
	        myWorkBook.write(os);
	        System.out.println("Done");
	    }
	}
