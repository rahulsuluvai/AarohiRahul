package org.sample;


import java.io.File;
import java.io.FileInputStream;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SampleBrowser {

	public static void main(String[] args) throws IOException {
		
		File file = new File("C:\\Users\\Lenovo\\eclipse-workspace\\FrameworkAugest9am\\TestDta\\Framework.xlsx");
		FileInputStream fileInputStream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(fileInputStream);
		Sheet sheet = workbook.getSheet("Sheet1");
 
		for (int i=0; i< sheet.getPhysicalNumberOfRows(); i++) 
		{
			Row row = sheet.getRow(i);
		for (int j=0; j< row.getPhysicalNumberOfCells(); j++) 
		{
			Cell cell = row.getCell(j);
			//String ss = cell.getStringCellValue();
			//System.out.println(ss);
			int cellType = cell.getCellType();
		if(cellType==1)
		{
			String ss = cell.getStringCellValue();
			System.out.println(ss);
		}
		
		else if(DateUtil.isCellDateFormatted(cell))
		{
			Date dd = cell.getDateCellValue();
			SimpleDateFormat dateFormat = new SimpleDateFormat("MMM-dd-yyyy");
			String format = dateFormat.format(dd);
			System.out.println(format);
			
		}
		else 
		{
			double numericCellValue = cell.getNumericCellValue();
			long l = (long) numericCellValue;
			String valueOf = String.valueOf(l);
			System.out.println(valueOf);
		}
		}
		}
		}
	
	}


