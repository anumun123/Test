package excel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;



public class ExcelRead {
		
		XSSFSheet sheet;

		ExcelRead() throws IOException
		{
			FileInputStream f = new FileInputStream("C:\\Users\\ANU R\\Desktop\\Worksheet.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook(f);
			sheet = wb.getSheet("Sheet1");
				
		}
		
		public String readData(int i,int j)
		{
			Row row = sheet.getRow(i);
			Cell cell = row.getCell(j);
			CellType type = cell.getCellType();
			
			switch(type)
			{
				case NUMERIC: return String.valueOf(cell.getNumericCellValue());
							
				case STRING: return String.valueOf(cell.getStringCellValue());
								
			}
			return null;
		}
	
	public int rowsize() 
	{
		return sheet.getLastRowNum() + 1;
	}
	
	
}


