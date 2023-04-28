package com.sgtesting.excelpart2assignments;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteToOtherExcelFile {

	public static void main(String[] args) {
		writeToDifferentExcelFile();
	}
	private static void writeToDifferentExcelFile()
	{
		FileInputStream fin=null;
		FileOutputStream fout=null;
		Workbook wb1=null;
		Workbook wb2=null;
		Sheet sh1=null;
		Sheet sh2=null;
		Row rowsh1=null;
		Row rowsh2=null;
		Cell cellsh1=null;
		Cell cellsh2=null; 
		try {
			fin=new FileInputStream("D:\\EXCEL\\Excel2\\WriteToOtherExcelFile.xlsx");
			wb1=new XSSFWorkbook(fin);
			
			
			
			
			fout=new FileOutputStream("D:\\EXCEL\\Excel2\\WriteToOtherExcelFile.xlsx");
			wb2.write(fout);
		}catch (Exception e) {
			e.printStackTrace();
		}

		finally
		{
			try {
					fin.close();
					fout.close();
					wb1.close();
					wb2.close();
			}catch (Exception e) {
				e.printStackTrace();
			}
		}
	}
}
