package xlsxDataHandling;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

public class ReadAndWriteXlsxData extends ReadExcelFile{
	static WebDriver driver;
	static String url;
	static String firstname;
	static String lastname;
	static String email;
	static String expectederror;
	
	@AfterMethod
	public static void teardown() {
		driver.close();
	}
	
	@BeforeMethod
	public static void calldriver() {
		System.setProperty("webdriver.chrome.driver", "C:\\Softwares\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");
		 driver = new ChromeDriver();
		
		driver.manage().window().maximize();
	}

	@Test(description="Read and write test data in xlsx format")
	public static void tC01() throws Exception {
		 File src = new File (System.getProperty("user.dir")+ "\\Resources\\resultSheet.xlsx");
		 FileInputStream inp = new FileInputStream(src);
		 XSSFWorkbook workbook = new XSSFWorkbook(inp);
		 System.out.println(src); 
	XSSFSheet sheet= workbook.getSheet("Sheet1");
	int totalrow =sheet.getLastRowNum();
	short clmcount =sheet.getRow(0).getLastCellNum();
	System.out.println(clmcount);
	for (int i =1; i<=totalrow; i++) {
		//for(int j=0; j<=clmcount;j++) {
		if(	sheet.getRow(i).getCell(0).getCellType() == CellType.STRING) {
		url =	sheet.getRow(i).getCell(0).toString();
			driver.navigate().to(url);
		}
			 if(sheet.getRow(i).getCell(1).getCellType() == CellType.STRING) {
				 firstname = sheet.getRow(i).getCell(1).toString();
				 WebElement firstnameField = driver.findElement(By.id("firstName"));
				 firstnameField.sendKeys(firstname);
			 }
			 if(sheet.getRow(i).getCell(2).getCellType() == CellType.STRING) {
				 lastname = sheet.getRow(i).getCell(2).toString();
				 WebElement lastnameField = driver.findElement(By.id("lastName"));
				 lastnameField.sendKeys(lastname);
			 }
			 if(sheet.getRow(i).getCell(3).getCellType() == CellType.STRING) {
				 email = sheet.getRow(i).getCell(3).toString();
				 WebElement emailField = driver.findElement(By.id("email"));
				 emailField.sendKeys(email);
			 }
			 WebElement submit = driver.findElement(By.id("nextBtn"));
			 submit.click();
			 
			 if(sheet.getRow(i).getCell(4).getCellType() == CellType.STRING) {
				 expectederror = sheet.getRow(i).getCell(4).toString();
				 System.out.println(expectederror);}
					
					 WebElement errorField = driver.findElement(By.xpath("//div[@class='mtop10'][2]/div/p"));
				String	 errormsg = errorField.getText();
					 System.out.println(errormsg);
				 boolean assertionflag = expectederror.equalsIgnoreCase(errormsg) ?  true :false ;
				 System.out.println(assertionflag);
			 
				 FileOutputStream out = new FileOutputStream(src)	;
				 if(assertionflag) {
					 sheet.getRow(i).createCell(5).setCellValue("Passed");
				 }
				 else { sheet.getRow(i).createCell(5).setCellValue("failed");}
				 workbook.write(out);
				 workbook.close();
				 out.close();
		
	}
	
	/*
	 * //if(sheet.getRow(1).getCell(0).getCellType()) XSSFCell url
	 * =sheet.getRow(1).getCell(0); System.out.println(url);
	 * driver.navigate().to(url);
	 */
    
		
	}
	
	/*
	 * public static void main(String[] args) throws Exception{ calldriver();
	 * tC01(); }
	 */
}
