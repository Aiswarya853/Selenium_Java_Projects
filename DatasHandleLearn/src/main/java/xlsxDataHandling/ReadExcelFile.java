package xlsxDataHandling;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelFile {
	
static String src =  System.getProperty("user.dir")+ "\\Resources\\resultSheet.xlsx";
public static XSSFWorkbook workbook;

public static void  readExcel() throws Exception {
FileInputStream inp = new FileInputStream(src);
XSSFWorkbook workbook = new XSSFWorkbook(inp);
System.out.println(src);




/*
 * public static void main(String[] args){ System.out.println(src); }
 */
}
}