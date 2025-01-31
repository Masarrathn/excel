package ExcelSheet;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Worksheet1 {

	public static  void main(String[] args) throws Exception {
				
				String excelFilePath= ".Resource\\Data.xlsx";
				
				FileInputStream inputstream = new FileInputStream(excelFilePath);
			    XSSFWorkbook workbook = new XSSFWorkbook(inputstream);
			    XSSFSheet sheet= workbook.getSheet("Sheet1");
				
				int rows=sheet.getLastRowNum();
				int cols=sheet.getRow(1).getLastCellNum();
				System.out.println("Number of Rows=" +rows);
				System.out.println("Number of Columns=" +cols);
				
				

			}

	}


