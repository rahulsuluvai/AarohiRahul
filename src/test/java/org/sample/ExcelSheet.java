package org.sample;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class ExcelSheet {

	public static void main(String[] args) throws IOException {
	WebDriverManager.chromedriver().setup();
	WebDriver driver = new ChromeDriver();
	driver.manage().window().maximize();
	driver.get("https://www.amazon.com/");
		
	WebElement search = driver.findElement(By.id("twotabsearchtextbox"));
	search.sendKeys("Iphone 11 pro  max");
	WebElement prd = driver.findElement(By.id("nav-search-submit-button"));
	prd.click();
	WebElement prntxt = driver.findElement(By.xpath("//span[contains(text(),'Apple iPhone 11 Pro Max, 64GB, Space Gray - Unlocked (Renewed Premium)')]"));
	String s = prntxt.getText();
	System.out.println(s);
	File file = new File("C:\\Users\\Lenovo\\eclipse-workspace\\FrameworkAugest9am\\TestDta\\Sheet2.xlsx");
	FileInputStream fis = new FileInputStream(file);
	Workbook wrk = new XSSFWorkbook();
	Sheet createSheet = wrk.createSheet("Sheet2");
	Row createRow = createSheet.createRow(0);
	Cell createCell = createRow.createCell(0);
	createCell.setCellValue(s);
	FileOutputStream fstr = new FileOutputStream(file);
	wrk.write(fstr);
	
	
	
	
	
	
	
	
	
	
	}

}
