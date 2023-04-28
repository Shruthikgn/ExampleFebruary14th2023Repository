package com.sgtesting.excelpart2assignments;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FruitNamesInto10thRow {

	public static void main(String[] args) {
		writeFruitTo10thRow();
	}

	private static void writeFruitTo10thRow()
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

		try
		{
			fin=new FileInputStream("D:\\EXCEL\\Excel2\\FruitNamesInto10thRow.xlsx");
			wb=new XSSFWorkbook(fin);
			
			sh1=wb.getSheet("Sheet1");
			sh2=wb.getSheet("Sheet2");
			if(sh2==null)
			{
				sh2=wb.createSheet("Sheet2");
			}
			rowsh1=sh1.getRow(9);
			rowsh2=sh2.getRow(9);
			if(rowsh2==null)
			{
				rowsh2=sh2.createRow(9);
			}
			int cc=rowsh1.getPhysicalNumberOfCells();
			for(int c=0;c<cc;c++)
			{
				cellsh1=rowsh1.createCell(c);
				cellsh2=rowsh2.createCell(c);
				if(cellsh2==null)
				{
					cellsh2=rowsh2.createCell(c);
				}
				String data=cellsh1.getStringCellValue();
				cellsh2.setCellValue(data);
			}				
			fout=new FileOutputStream("D:\\EXCEL\\Excel2\\FruitNamesInto10thRow.xlsx");
			wb.write(fout);
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
