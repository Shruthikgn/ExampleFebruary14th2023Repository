package com.sgtesting.excelpoiassignments;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/* Programatically write 20 fruit names 
 * in 1st sheet, 1st column of an excel file
 */
public class FruitNames {

	public static void main(String[] args) {
	writeFruitNames();	

	}

	private static void writeFruitNames()
	{
		FileOutputStream fout=null;
		Workbook wb=null;
		Sheet sh=null;
		Row row =null;
		Cell cell=null;
		
		try
		{
			wb=new XSSFWorkbook();
			sh=wb.createSheet();
			for(int i=0;i<20;i++)
			{
				row=sh.createRow(i);
				cell=row.createCell(0);
				
				cell.setCellValue("Fruit"+(i+1));
				
				fout= new FileOutputStream("D:\\EXCEL\\FruitsList.xlsx");
				wb.write(fout);
			}
		}catch (Exception e) {
			e.printStackTrace();
		}
		finally
		{
			try {
				fout.close();
				wb.close();
				
			}catch (Exception e) {
				e.printStackTrace();
			}
		}
	}
}
