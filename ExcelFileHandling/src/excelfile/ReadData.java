package excelfile;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

@SuppressWarnings("unused")
public class ReadData 
{
	@SuppressWarnings("resource")
	@Test
	public void readdata() throws IOException
	{
		FileInputStream fis = new FileInputStream("./testdata/Readwrite.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		/*
		 * Sheet s = wb.getSheet("Blackcap"); Row r = s.getRow(15); Cell c =
		 * r.getCell(2); System.out.println(c.getStringCellValue());
		 */
		//System.out.println(wb.getSheet("Blackcap").getRow(15).getCell(2).getStringCellValue());
		System.out.println(wb.getSheet("Blackcap").getRow(19).getCell(4).toString());
	}
}











