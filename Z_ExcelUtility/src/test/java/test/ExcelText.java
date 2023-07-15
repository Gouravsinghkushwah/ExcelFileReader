 package test;
import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
public class ExcelText {

	static String url="jdbc:oracle:thin:@localhost:1521:orcl"; //--- Give here your local Oracle Database URL.
	static String username = "gourav";  //---- Give here your Oracle UserName.
	static String password = "gourav";   //----- Give here your Oracle Password.
	private static Connection con = null;
	@SuppressWarnings("static-access") 
	public static void main(String[] args) throws Exception
	{
		//---- Useful Comments:
		// 1> File Name : StudentData.xlsx,      sheetName : sheet1
		// 2> File Name : Book1.xlsx             sheetName : sheet1  (Empty file)
		// 2> File Name : Book2.xlsx,            sheetName : Index
		
		
		
		String filepath="./data/StudentData.XLSX"; //----- Give here your Local system excel file(must be with .xlsx extension) complete URL.
		String sheetName = "sheet1";   //------- Give here file sheetName.
		
		if(!(filepath.substring(filepath.length()-5)).equalsIgnoreCase(".xlsx"))
		{
			System.out.println("Only .xlsx formate accepted.");
			System.exit(0);
		}
		
		File myFile = new File("./data/Book2.xlsx");
		Workbook wb = WorkbookFactory.create(myFile);
		List<String> sheetNames = new ArrayList<String>();

		for (int i=0; i<wb.getNumberOfSheets(); i++) {
		    sheetNames.add( wb.getSheetName(i) );
		    
		}
		System.out.println("sheetName : " +sheetName);
		try
		{
			
			Class.forName("oracle.jdbc.driver.OracleDriver");
			System.out.println("Driver is loaded successfully");
			con = DriverManager.getConnection(url,username,password);
			System.out.println("Connectio stablished is successfull"+con);
		}
		catch(Exception e)
		{
			System.out.println(e.getMessage());
			e.printStackTrace();
		} 
		
		ExcelFile ef = new ExcelFile(filepath,sheetName);
		ef.getRowColCount();
 		ef.getCellData();	
	}
	
	public static Connection getCon()
	{
		return con;
	}
	
}
