package com.abc.Magentoscript;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class MagentoTest 
{
	public static XSSFSheet sheet;
	public static WebDriver driver;
	public static String getCellValue(int r, int c) 
	{
		XSSFRow row=sheet.getRow(r);
		XSSFCell cell=row.getCell(c);
		String text=cell.getStringCellValue();
		return text;
		}

	public static void main(String[] args) throws IOException 
	{
		FileInputStream fis=new FileInputStream("C:\\Users\\user\\Desktop\\ABC\\ExcelWorkBook\\HybirdDrivenData.xlsx");
		
		XSSFWorkbook book=new XSSFWorkbook(fis);
		sheet = book.getSheetAt(0);
		int rowno = sheet.getPhysicalNumberOfRows();
		System.out.println("the no of rows are: " +rowno);
		
		for(int i=1; i<rowno; i++)
		{
			String action=getCellValue(i,2);
			System.out.println(action);
			
			switch(action)
			{
			case "open":
			driver = new ChromeDriver();
			driver.manage().window().maximize();
			driver.manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
			driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
			break;
			
			case "navigate":
			driver.get(getCellValue(i,4));
			break;
			
			case "Click":
			driver.findElement(By.xpath(getCellValue(i, 3))).click();
			break;
			
			case "type":
			driver.findElement(By.xpath(getCellValue(i, 3))).sendKeys(getCellValue(i, 4));
			break;
			
			case "close":
			driver.close();
			break;
			
			}
		}
	}
}
