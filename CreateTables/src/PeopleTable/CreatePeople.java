package PeopleTable;
import java.io.FileInputStream;
import java.sql.*;
import java.util.Random;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class CreatePeople 
{
	private static String url ;		//Type url
	private static String user ;	//Type username
	private static String pass ;	//Type password
	private static String fileLocation ;	//Type file location
	
	public static int numberOfUser = 300;
	
	private static Connection cn = null;
	private static PreparedStatement pstmt = null;
	
	public static void main(String[] args) throws Exception
	{
		//Variable and classes declaration:
		Cell col1 = null;
		Cell col2 = null;
		Cell col3 = null;
		String name, lastName, phoneNumber, email, query;
		
		//Create InputStram for excel file and set up Apache POI for reading
		FileInputStream f = new FileInputStream(fileLocation);		
		Workbook workbook = new WorkbookFactory().create(f);	
		Sheet sheet = workbook.getSheet("Sheet1");
		DataFormatter dataFormater = new DataFormatter();
		
		//Crate connection to MySQL Database
		createConnection();
		
		//Prepare random generator and variables
		int nameRD, lastNameRD,  ageRD;
		Random rd = new Random();
		
		int affectedRow, phoneNumberIndex = 0;

		int counter = 0;
		
		while(counter <= numberOfUser)
		{	
			//Generate random indexes to retrieve data from excel sheet
			nameRD = rd.nextInt((164052 - 0) + 1) + 0;
			lastNameRD = rd.nextInt((98342 - 0) + 1) + 0;
			ageRD = rd.nextInt((80 - 15) + 1) + 15;
			
			//Retrieve data from excel sheet
			col1 = sheet.getRow(nameRD).getCell(0);
			col2 = sheet.getRow(lastNameRD).getCell(1);
			col3 = sheet.getRow(phoneNumberIndex).getCell(2);
			
			//Covert cell values to string
			name = dataFormater.formatCellValue(col1);
			lastName = dataFormater.formatCellValue(col2);
			phoneNumber = dataFormater.formatCellValue(col3);	
			email = name + lastName + "@email.com";
			
			//Prepare query and execute
			query = "INSERT INTO person(Last_Name, First_Name, Email, Phone, Age) values(?,?,?,?,?)";
			pstmt = cn.prepareStatement(query);
			pstmt.setString(1, lastName);
			pstmt.setString(2, name);
			pstmt.setString(3, email);
			pstmt.setString(4, phoneNumber);
			pstmt.setInt(5, ageRD);
			
			affectedRow = pstmt.executeUpdate();		//Excute Query
		
			//System.out.println(affectedRow);		//Print row affected	
			
			//Increase counter and phoneNumber index
			counter++;
			phoneNumberIndex++;
			
		}
		cn.close();									//Close connection
		System.out.println("Done");
	}
	
	private static void createConnection()
	{
		try
		{
			cn = DriverManager.getConnection(url, user, pass);					
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}
}
