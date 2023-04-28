package com.sgtesting.excelpoiassignments;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FlowerNames10thRow {

	public static void main(String[] args) {
		writeFlowerNames();
		}
	private static void writeFlowerNames()
	{
		FileOutputStream fout=null;
		Workbook wb=null;
		Sheet sh=null;
		Row row=null;
		Cell cell=null;
		try {
			wb=new XSSFWorkbook();
			sh=wb.createSheet();
			for(int rw=9;rw<29;rw++)
			{
				row=sh.createRow(rw);
				cell=row.createCell(0);
				cell.setCellValue("Flower"+(rw-8));
			}
			fout=new FileOutputStream("D:\\EXCEL\\FlowerNames.xlsx");
			wb.write(fout);
		}catch (Exception e) {
			e.printStackTrace();
		}
		finally
		{
			try {
				
			}catch (Exception e) {
				e.printStackTrace();
			}
		}
	}
	
}
