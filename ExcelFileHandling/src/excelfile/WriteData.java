package excelfile;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class WriteData 
{
	@SuppressWarnings("resource")
	@Test
	public void insertdata() throws IOException
	{
		FileInputStream fis = new FileInputStream("./testdata/Readwrite.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		//Sheet s = wb.createSheet("Wagner");
		//Row r = s.createRow(7);
		//Cell c = r.createCell(5);
		//c.setCellValue("WTC Final");
		//FileOutputStream fos = new FileOutputStream("./testdata/Readwrite.xlsx");
		//wb.write(fos);
		Sheet s = wb.getSheet("Blackcap");
		Row r = s.getRow(14);
		Cell c = r.createCell(7);
		c.setCellValue("Southampton");
		FileOutputStream fos = new FileOutputStream("fis");
		wb.write(fos);
	}
}














