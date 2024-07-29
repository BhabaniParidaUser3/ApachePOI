package excelOperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcelUsingForLoop {

	public static void main(String[] args) throws IOException {

		//create workbook->inside workbook create sheet->inside sheet create multiple rows->each row having multiple cells
		
		//Create a empty workbook
		XSSFWorkbook workbook=new XSSFWorkbook();
		
		//Inside workbook create a new sheet
		XSSFSheet sheet=workbook.createSheet("Emp Info");
		
		//have some way to hold data using data structure concept(Object Array/ArrayList/hashMap)
		
		//create 2 dimensional Object array(which can hold heterogeneous data )In excel file we have row & column ,so result will in 2Dimensional 
		
	 Object[][]	empdata= {{"EmpID","Name","Job"},{101,"David","Engineer"},{102,"Smith","Manager"},{103,"Scott","Analyst"}};
	 
	 //write the empdata into sheet
	 //using normal For loop
	 //we have to think how many rows and how many columns will be from 2d array
	
	 //get number of rows
	 int rows= empdata.length;
	 
	 //get number of cols
	 int cols=empdata[0].length;
	 
	 System.out.println(rows);
	 System.out.println(cols);
	 
	 //outer for loop is for rows
	 for(int r=0;r<rows;r++)
	 {
		 //so r=0,in excel sheet we have to create a new row & then we can write multiple cell in that particular row
		 //inner for loop is responsible for writing the cell
		 //so before that we have to create a row
		 
		//here create row using sheet object
		 XSSFRow row=sheet.createRow(r);
		 //inner for loop is for cols
		 for(int c=0;c<cols;c++)
		 {
			 //once row created,we have to create cell using row object
			 XSSFCell cell=row.createCell(c);
			 //now once cell created we have to write the data inside the cell from capturing data from ObjectArray
			 Object value=empdata[r][c];
			 
			 //now store/update the reference value into excel sheet
			 //before updating we have to check what kind of value it is?(if it is a string /integer/boolean/)
			 //now check the value
			 //value is in form of object,we don't know its a string/boolean/integer
			 if(value instanceof String)
			 {
				 cell.setCellValue((String) value);
			 }
			 if(value instanceof Integer)
			 {
				 cell.setCellValue((Integer) value);
			 }
			 if(value instanceof Boolean)
			 {
				 cell.setCellValue((Boolean) value);
			 }
			 
		 }
	 }
	 
	 //where the file need to create mention that path
	 String filepath=".\\dataFiles\\Employees.xlsx";
	 
	 //want to open the file fileOutput stream mode because we are going to write the data
	 FileOutputStream outputStream=new FileOutputStream(filepath);
	 
	 
	 //now write the workbook into the ExcelFile
	 workbook.write(outputStream);
	 
	 //close outputStream
	 outputStream.close();
	 
	 System.out.println("Employees.xlsx File created Successfully");
	 
	 
	 
	}

}
