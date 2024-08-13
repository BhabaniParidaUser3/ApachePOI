package excelOperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataBaseToExcel {

	public static void main(String[] args) throws SQLException, IOException {

		//Connect to database
	Connection conn=	DriverManager.getConnection("jdbc:mysql://localhost:3306/world","root","ABcd12@@##%%");
	
	//Statement or query
	Statement stmt=conn.createStatement();
	ResultSet rs=stmt.executeQuery("Select * from city;");
	
	//Excel
	XSSFWorkbook workbook=new XSSFWorkbook();
	XSSFSheet sheet=workbook.createSheet("City Data");
	
	XSSFRow row=sheet.createRow(0);
	row.createCell(0).setCellValue("ID");
	row.createCell(1).setCellValue("Name");
	row.createCell(2).setCellValue("CountryCode");
	row.createCell(3).setCellValue("District");
	row.createCell(4).setCellValue("Population");
	
	int r=1;
	while(rs.next())
	{
		int id=rs.getInt("ID");
		String name=rs.getString("Name");
		String countryCode=rs.getString("CountryCode");
		String district=rs.getString("District");
		int population=rs.getInt("Population");
		row=sheet.createRow(r++);
		
		row.createCell(0).setCellValue(id);
		row.createCell(1).setCellValue(name);
		row.createCell(2).setCellValue(countryCode);
		row.createCell(3).setCellValue(district);
		row.createCell(4).setCellValue(population);
	}
	
	String outputpath=".\\dataFiles\\City3.xlsx";
	FileOutputStream outputstream=new FileOutputStream(outputpath);
	workbook.write(outputstream);
	workbook.close();
	outputstream.close();
	conn.close();
	System.out.println("Done");
	
	



	
	
	
	}

}
