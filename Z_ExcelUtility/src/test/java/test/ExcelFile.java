package test;
import java.beans.Statement;

import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFile {
	static XSSFWorkbook wb;
	static XSSFSheet sheet;
	static String filePath;
	static int RowCount,Coulcount;
	static Object arr[] = new Object[10];
	static int count=0;
	static String flag="N";
	ExcelFile(){}
	static Scanner sc = new Scanner(System.in);
	
	
 	//---- insertData() Method interacting with database.
	public static void insertData(Object... obj)   
	{		
		if(Coulcount==4 || flag.equalsIgnoreCase("Y")) {
		try
		{
			
			Connection con =  ExcelText.getCon();
 			PreparedStatement ps = con.prepareStatement
			("insert into Org values(?,?,?,?)");
 			ps.setObject(1,obj[0]);
			ps.setObject(2,obj[1]);
			ps.setObject(3,obj[2]);
			ps.setObject(4,obj[3]);
			count+= ps.executeUpdate();
//			System.out.println("count - -"+count);  // To check executeUpdate() is working with all rows or not.
			
		} catch(Exception e)
		{
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
		}  
	}
	public ExcelFile(String filePath, String sheetName)
	{
		try
		{
		  wb = new  XSSFWorkbook(filePath);
		  sheet = wb.getSheet(sheetName); 
		  
		}
		catch(Exception e)
		{
			System.out.println(e.getCause());
			System.out.println(e.getMessage());
			e.printStackTrace();		}
	}
	
	public static void getCellData()
	{
		try
		{
			for(int i=0;i<RowCount;i++)
			{
			for(int j=0;j<Coulcount;j++)
			{
 				DataFormatter formatter = new DataFormatter();
				Object value = formatter.formatCellValue(sheet.getRow(i).getCell(j));
				arr[j] = value;	 
			//String value = 	sheet.getRow(i).getCell(j).getStringCellValue();
				
				System.out.print(String.format("%-18s", arr[j]));
			}
			if(i>0)
			insertData(arr);
			System.out.println();
		}
			
		} catch(Exception e)
		{
			e.printStackTrace();	
		}	
		if(Coulcount==4 || flag.equalsIgnoreCase("Y"))
			System.out.println("Database Upadted successfully");
		else if( flag.equalsIgnoreCase("n") || Coulcount!=4) 
			System.err.println("\n\nDatabase Not Upadted.");
		  
	}
	public static void getRowColCount()
	{
		try
		{
  		  RowCount  = sheet.getPhysicalNumberOfRows();
		  Coulcount  = sheet.getRow(0).getLastCellNum();
		  System.out.println("Number of Columns : "+Coulcount);
		System.out.println("Number of rows : "+RowCount);
		if(RowCount==1)
		{
			System.out.println("Fount Excel File with no data."
					+ "\ncan't update Database."
					+ "\n Please try another file.");
			System.exit(0);
		}
		if(Coulcount!=4) {
 		System.out.println("Your Excel Sheet has not 4 column."
 				+ "\nStill you want to update Database Enter Y otherwise enter N.");
		flag = sc.nextLine();
 
		}
		}
		catch(Exception e)
		{
			System.out.println(e.getCause());
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
		 
	}
	
}
