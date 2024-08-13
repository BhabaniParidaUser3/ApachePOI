package excelOperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataFromHashMapToExcel {

	public static void main(String[] args) throws IOException {

		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet();
		
		Map<String,String> data=new HashMap<String,String>();
		data.put("101", "Raj");
		data.put("102", "Bhabani");
		data.put("103", "Tiki");
		data.put("104", "manu");
		data.put("105", "Pratish");
		
		int rowno=0;
		for(Map.Entry entry:data.entrySet()) {
			XSSFRow row=sheet.createRow(rowno++);
			
			row.createCell(0).setCellValue((String)entry.getKey());
			row.createCell(1).setCellValue((String)entry.getValue());

			
		}
		String path=".//dataFiles//Student.xlsx";
		FileOutputStream outputStream=new FileOutputStream(path);
		workbook.write(outputStream);
		outputStream.close();
		System.out.println("Student1.xlsx created succesfully");
		

		
	}

}
