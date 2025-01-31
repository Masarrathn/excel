package ExcelSheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;


import org.apache.poi.ss.usermodel.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {

	public static void main(String[] args) {
		
		String filePath = ".\\Resource1\\Data.xlsx";
			try {
				
			FileInputStream file= new FileInputStream(new File(".\\Resource1\\Data.xlsx"));
			Workbook workbook =  new XSSFWorkbook(file);
			Sheet sheet= workbook.getSheetAt(0);
			
			for (Row row : sheet) {
				
				for (Cell cell : row) {
					
					switch (cell.getCellType()) {
					
					case STRING : 
						System.out.print(cell.getStringCellValue() + "\t");
                        break;
                        
					case NUMERIC:
                        System.out.print(cell.getNumericCellValue() + "\t");
                        break;
					}
				}
				
				System.out.println();
			}
			
				workbook.close();
	            file.close(); 
			}
					 
	                        
			 catch (IOException e) {
				
			
	            e.printStackTrace(); 
	            
			 
			 }
	}
}
	
			

			
			
	
	
		
		
					 
					 
					 
					 
	       
				 
			
			
		


