package com.sgtesting.excelpart2assignments;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FlowersRowsToColumns {

	public static void main(String[] args) {
		rowToColumn();
	}
	private static void rowToColumn()
	{
		FileInputStream fin=null;
		FileOutputStream fout=null;
		Workbook wb=null;
		Sheet sh1=null;
		Sheet sh2=null;
		Row rowsh1=null;
		Row rowsh2=null;
		Cell cellsh1=null;
		Cell cellsh2=null; 
		
		try {
			fin=new FileInputStream("D:\\EXCEL\\FlowersRowsToColumns.xlsx");
			wb=new XSSFWorkbook(fin);
			sh1=wb.getSheet("Sheet1");
			sh2=wb.getSheet("Sheet2");
			if(sh2==null)
			{
				sh2=wb.createSheet("Sheet2");
			}
			int rc=sh1.getPhysicalNumberOfRows();
			int cc=rowsh1.getPhysicalNumberOfCells();
			for(int r=0;r<cc;r++)
			{
				
			
			
			
			
			
					String data=cellsh1.getStringCellValue();
					cellsh2.setCellValue(data);
			
			}
			fout=new FileOutputStream("D:\\EXCEL\\FlowersRowsToColumns.xlsx");
			wb.write(fout);
		
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		finally
		{
			try {
				fin.close();
				fout.close();
				wb.close();
			}catch (Exception e) {
				e.printStackTrace();
			}
		}
	}
}