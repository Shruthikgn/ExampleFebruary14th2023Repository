package com.sgtesting.excelpoiassignments;
/* Programmatically write 20 state names in
 *  1st sheet 5th column of an excel file
 */
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class StateNames5thColumn {

	public static void main(String[] args) {
		writeStateNames();

	}
	private static void writeStateNames()
	{
		FileOutputStream fout=null;
		Workbook wb=null;
		Sheet sh=null;
		Row row=null;
		Cell cell=null;
		try {
			wb=new XSSFWorkbook();
			sh=wb.createSheet();
			for(int rw=0;rw<20;rw++)
			{
				row=sh.createRow(rw);
				cell=row.createCell(4);
				cell.setCellValue("State"+(rw+1));
			}
			fout=new FileOutputStream("D:\\EXCEL\\Statenames.xlsx");
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
