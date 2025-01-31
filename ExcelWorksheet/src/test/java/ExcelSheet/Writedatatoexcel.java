package ExcelSheet;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Writedatatoexcel {

	public static void main(String[] args) {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Info");
		
		Object empdata[][] = { {"Name", "EmpAge", "Email"}, {"Emily", "26", "emily@test.com"}, {"Richard", "32", "richard@example.com"}
				,{"Virat" ,"34" , "Virat@exampe.com"}};
				
		int rows=empdata.length;
		int cols=empdata[0].length;
		
		System.out.println(rows);
		System.out.println(cols);
		

	}

}
