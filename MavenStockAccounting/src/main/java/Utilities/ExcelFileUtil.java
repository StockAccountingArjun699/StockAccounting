package Utilities;



import java.io.FileInputStream;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtil 
{
Workbook wb;
public Object rowCount;

//it will load Excel Sheet
public ExcelFileUtil() throws Throwable
{
	FileInputStream fis=new FileInputStream("C:\\Users\\nagarjuna.y\\workspace\\StockAccounting\\MavenStockAccounting\\Testinputs\\Inputsheet (2).xlsx");
	 wb =WorkbookFactory.create(fis);
}

// row count

public int rowCount(String sheetname)
{
	return wb.getSheet(sheetname).getLastRowNum();
}
//column count

public int colCount(String sheetname,int row)
{
	return wb.getSheet(sheetname).getRow(row).getLastCellNum();
}
//reading the data
public String getData(String sheetname,int row,int column)
{
	String data="";
	
	{
		int celldata=(int)wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();
	data=String.valueOf(celldata);
	}
	{
	data=wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();
	}
	
	return data;
}

//Storing data into Excel Sheet Pass or Fail and Not Executed
public void setData(String sheetname,int row,int column,String status) throws Throwable
{
	Sheet sh=wb.getSheet(sheetname);
	Row rownum=sh.getRow(row);
	Cell cell=rownum.createCell(column);
	cell.setCellValue(status);
	
	if(status.equalsIgnoreCase("Pass"))
	{
		//Create Cell Style
		CellStyle style=wb.createCellStyle();
		//Craete Font
		Font font=wb.createFont();
		//Apply Colour To Text
		font.setColor(IndexedColors.GREEN.index);
		//Apply Bold To Text
		font.setBold(true);
		//Set Font
		style.setFont(font);
		//Set Cell Style
		rownum.getCell(column).setCellStyle(style);
	}else
		if(status.equalsIgnoreCase("Fail"))
		{
			//Create Cell Style
			CellStyle style=wb.createCellStyle();
			//Craete Font
			Font font=wb.createFont();
			//Apply colour to text
			font.setColor(IndexedColors.RED.index);
			//Apply Bold to text
			font.setBold(true);
			//Set Font
			style.setFont(font);
			//Set Cell Style
			rownum.getCell(column).setCellStyle(style);
		}else
			if(status.equalsIgnoreCase("Not Executed"))
			{
				//Create Cell Style
				CellStyle style=wb.createCellStyle();
				//Craete Font
				Font font=wb.createFont();
				//Apply colour to text
				font.setColor(IndexedColors.BLUE.index);
				//Apply Bold to text
				font.setBold(true);
				//Set Font
				style.setFont(font);
				//Set Cell Style
				rownum.getCell(column).setCellStyle(style);
			}
	
	FileOutputStream fos=new FileOutputStream("C:\\Users\\nagarjuna.y\\workspace\\StockAccounting\\MavenStockAccounting\\testoutputs\\out putsheet.xlsx.xlsx");
	wb.write(fos);
	fos.close();
		
}

public static void main(String[] args) throws Throwable 
{
	ExcelFileUtil excel=new ExcelFileUtil();
	
	System.out.println(excel.rowCount("MasterTestCases"));
	
	System.out.println(excel.colCount("Sheet1", 1));
	
	System.out.println(excel.getData("Sheet1", 1, 1));
	
	
	excel.setData("Sheet1", 1, 2, "Pass");
	excel.setData("Sheet1", 2, 2, "Fail");
	excel.setData("Sheet1", 3, 2, "Not Executed");
}
}
