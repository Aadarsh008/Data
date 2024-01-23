package dataJD;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Configure {
	
 public static void createExcel(  ArrayList<String> name ) throws IOException
 { 
	 
	 XSSFWorkbook workbook = new XSSFWorkbook(); 
	 XSSFSheet spreadsheet = workbook.createSheet(" Customer Data "); 
	 
	 Row row = spreadsheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Name");
	int n=1;
	for (String name1 : name)
	{
		 row = spreadsheet.createRow(n++);
		cell = row.createCell(0);
		cell.setCellValue(name1);
	}
	 FileOutputStream out;
	try {
		out = new FileOutputStream(new File("C:\\New folder\\GFGsheet.xlsx"));
		workbook.write(out); 
        out.close(); 
	} catch (FileNotFoundException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	} 
 } 
 
public static void main(String[] args) throws Exception
    {
        String url= "jdbc:mysql://localhost:3306/employee";
        String username = "root"; 
        String password = "";
        String query
            = "select *from customer"; 
        Class.forName("com.mysql.cj.jdbc.Driver");
        Connection con = DriverManager.getConnection(
                url, username, password);
            System.out.println(
                "Connection Established successfully");
            Statement st = con.createStatement();
            ResultSet rs
                = st.executeQuery(query); 
            ArrayList<String> name = new ArrayList<String>();
            while(rs.next())
            {
             name.add(rs.getString("CustomerName")); // Retrieve name from db
            }
            System.out.println(name); 
            st.close(); // close statement
            con.close(); // close connection
            System.out.println("Connection Closed....");
            
           createExcel(name);
      }
	
}
