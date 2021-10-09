package writeFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//()   = 	 0   +
public class Write_File {
	public static void main(String[] args) throws IOException {
		File f   = new File("C:\\Users\\user\\eclipse-workspace\\DataDrivenProject\\DataDrivenPProect.xlsx");
		
		FileInputStream fis =new FileInputStream(f);
		
		Workbook wb = new XSSFWorkbook(fis);

		Sheet getsheet = wb.getSheet("DataSheet");
		
		Row createRow = getsheet.createRow(0);
		Cell createCell = createRow.createCell(0);
		createCell.setCellValue("Pooja");
		
		
		
		
		
		
		FileOutputStream fos =new FileOutputStream(f);
		
		wb.write(fos);
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
	}

}
