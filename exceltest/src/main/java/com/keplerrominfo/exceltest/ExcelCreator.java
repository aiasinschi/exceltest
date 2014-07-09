package com.keplerrominfo.exceltest;

import java.io.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCreator {

	public static void main(String[] args) {
		String fname = "workbook.xlsx";
		InputStream inp;
		Workbook wb;
		try {
			inp = new FileInputStream(fname);
			wb = WorkbookFactory.create(inp);

			/*
			 * Sheet sheet1 = wb.createSheet("Sheet 1"); Row row1 =
			 * sheet1.createRow((short) 0); Cell cell1 = row1.createCell((short)
			 * 0); cell1.setCellValue(0.1234);
			 */

			/*
			 * Get first sheet in document
			 */
			Row row1 = wb.getSheetAt(0).getRow(0);
			if (row1 == null)
				System.out.print("row is null");
			Cell cell1 = row1.getCell(0);
			if (cell1 == null) { 
				System.out.print("cell is null");
				cell1 = row1.createCell(0);
			}
			cell1.setCellValue(0.111);

			/*
			 * Force recalculation of second sheet with formula
			 */
			Sheet Sheet2 = wb.getSheetAt(1);
			Sheet2.setForceFormulaRecalculation(true);
			
			FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
			wb.write(fileOut);
			fileOut.close();
			System.out.println("File written.");
		} catch (Exception ex) {
			ex.printStackTrace();
		}

	}

}
