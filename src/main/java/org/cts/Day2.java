package org.cts;

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

public class Day2 {

	
	public static Date dateCellValue;

	public static void main(String[] args) throws IOException {
		File f=new File("C:\\Users\\Sridharannamalai\\eclipse-workspace\\Test\\Data\\Book1.xlsx");
	FileInputStream stream =new FileInputStream(f);
	Workbook w=new XSSFWorkbook(stream);
	Sheet s=w.getSheet("Sheet1");
	Row r=s.getRow(2);
	Cell c=r.getCell(4);
	System.out.println(c);
	int celltype =c.getCellType();
	System.out.println(celltype);
	if(celltype==1) {
		String stringCellValue = c.getStringCellValue();
	System.out.println(stringCellValue);
	}
	else if(celltype==0) {
		if(DateUtil.isCellDateFormatted(c)) {
		Date dateCellValue = c.getDateCellValue();
	System.out.println(dateCellValue);
	}
	SimpleDateFormat s1=new SimpleDateFormat("dd-MM-yyy");
		String format=s1.format(dateCellValue);
	System.out.println(format);
	}
	else
	{
	double numericCellValue=c.getNumericCellValue();
	long l=(long)numericCellValue;
	String valueOf = String.valueOf(l);
	System.out.println(valueOf);
	}
	}
}
