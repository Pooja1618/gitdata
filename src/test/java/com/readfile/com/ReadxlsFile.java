package com.readfile.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//()   = 	 0   +
public class ReadxlsFile {
	
	public static void main(String[] args) throws IOException {
		File f =new File("C:\\Users\\user\\eclipse-workspace\\DataDrivenProject\\DataDrivenPProect.xlsx");
		
		FileInputStream fis=new FileInputStream(f);
		Workbook wb =new XSSFWorkbook(fis);
		
		Sheet sheetAt = wb.getSheetAt(0);
		
		int physicalNumberOfRows = sheetAt.getPhysicalNumberOfRows();

		
	for (int i = 0; i < physicalNumberOfRows; i++) {
	Row row = sheetAt.getRow(i);
	
	int physicalNumberOfCells = row.getPhysicalNumberOfCells();
	for (int j = 0; j < physicalNumberOfCells; j++) {
		Cell cell = row.getCell(j);
		
		CellType cellType = cell.getCellType();
		
	if (cellType.equals(cellType.NUMERIC)) {
	 
		double numericCellValue = cell.getNumericCellValue();
		
		int cellvalue = (int) numericCellValue;
		
		System.out.println(cellvalue);
		
		
	}
	else {
		System.out.println(cell);
	}
	wb.close();
	
	}
	
	}
		
	}
	

}


