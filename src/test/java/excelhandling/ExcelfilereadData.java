package excelhandling;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ExcelfilereadData {
	
	@Test
	public void test1() throws IOException {
		
		File f=new File("C:\\Users\\user\\Desktop\\Testdata.xlsx");
		
		FileInputStream fis=new FileInputStream(f);  // to read data
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheet("Data");
		XSSFRow row=sheet.getRow(0);
		XSSFCell cell=row.getCell(0);
		
		//for cell value
		String data=cell.getStringCellValue();
		System.out.println(data);
		
		//to get no.of rows
		int rowsCount=sheet.getLastRowNum();
		System.out.println(rowsCount);
		
		
		//to get no.of cols
		int colsCount=row.getLastCellNum();
		System.out.println(colsCount);
		
		
		//total data in excel
		for(int i=0;i<=rowsCount; i++){
			for(int j=0; j<colsCount; j++) {
				System.out.print(sheet.getRow(i).getCell(j).getStringCellValue()+  "  /");
			}
			
			System.out.println();
			
		}
	}

}
