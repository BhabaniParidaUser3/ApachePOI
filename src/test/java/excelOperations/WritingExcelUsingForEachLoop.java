package excelOperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcelUsingForEachLoop {

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
	 //using  For Each loop
	 int rowcount=0;
	 for(Object emp[]:empdata)
	 {
		 XSSFRow row=sheet.createRow(rowcount++);
		 int cellcount=0;
		 for(Object value:emp)
		 {
			 XSSFCell cell=row.createCell(cellcount++);
			 if(value instanceof String)
			 {
				 cell.setCellValue((String)value);
			 }
			 if(value instanceof Boolean)
			 {
				 cell.setCellValue((Boolean)value);
			 }
			 if(value instanceof Integer)
			 {
				 cell.setCellValue((Integer)value);
			 }
		 }
	 }
	 

	// where the file need to create mention that path
	String filepath = ".\\dataFiles\\Employees1.xlsx";

	// want to open the file fileOutput stream mode because we are going to write
	// the data
	FileOutputStream outputStream = new FileOutputStream(filepath);

	// now write the workbook into the ExcelFile
	workbook.write(outputStream);

	// close outputStream
	outputStream.close();

	System.out.println("Employees1.xlsx File created Successfully");

}

}
