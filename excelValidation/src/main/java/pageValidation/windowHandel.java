package pageValidation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class windowHandel {
	public File file;
	public FileInputStream fis;
	public FileOutputStream fos;
	public XSSFSheet sheet;
	public File file1;
	public XSSFWorkbook wb;
	
	public String excelfileread(String filename, int rown , int colnum,String sheetname) throws Exception{
		String data="";
		fis=new FileInputStream("C:\\Users\\niraj\\eclipse-workspace\\excelValidation\\excelInput\\Book1.xlsx");
		wb= new XSSFWorkbook(fis);
		sheet =wb.getSheet(sheetname);
		Row r= sheet.getRow(rown);
		Cell c= r.createCell(colnum);
		return data;
	}
	public void excelfilewrite(String filename, int rown, int colnum, String sheetname,String data) throws Exception {
		fis=new FileInputStream("C:\\Users\\niraj\\eclipse-workspace\\excelValidation\\excelInput\\Book1.xlsx");
		wb= new XSSFWorkbook(fis);
		sheet= wb.getSheet(sheetname);
		Row r= sheet.getRow(rown);
		Cell c= r.createCell(colnum);
		c.setCellValue(data);
		fos= new FileOutputStream("C:\\Users\\niraj\\eclipse-workspace\\excelValidation\\excelOutput\\output.xlsx");
		wb.write(fos);
	}
	
		

}
