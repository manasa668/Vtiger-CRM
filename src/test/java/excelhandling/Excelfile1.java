package excelhandling;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.testng.annotations.Test;

public class Excelfile1 {
	

	@Test
	public void test1() throws IOException {
		
		File f=new File("C:\\Users\\user\\Desktop\\Testdata.xlsx");
		
		FileInputStream fis=new FileInputStream(f);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		Syste.out.println("hello");
		XSSFSheet sheet=wb.getSheet("data1");
		WebDriver driver=new EdgeDriver();
		driver.get("https://www.gmail.com/");
		
	}
}
		
	
