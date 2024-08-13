package excelOperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataFromExcelToHashMap {

	public static void main(String[] args) throws IOException {

		String path=".\\dataFiles\\Student.xlsx";
		FileInputStream inputstream=new FileInputStream(path);
		XSSFWorkbook  workbook =new XSSFWorkbook(inputstream);
		XSSFSheet sheet=workbook.getSheet("Sheet0");
		int rows=sheet.getLastRowNum();
		int cols=sheet.getRow(0).getLastCellNum();
		
		HashMap<String,String> data=new HashMap<String,String>();
		//Reading data from Excel
		for(int r=0;r<=rows;r++)
		{
			String key=sheet.getRow(r).getCell(0).getStringCellValue();
			String value=sheet.getRow(r).getCell(1).getStringCellValue();
			data.put(key, value);
			
		}
		
		//Read data from hashMap
		
		for(Map.Entry entry:data.entrySet())
		{
			System.out.println(entry.getKey()+"  "+entry.getValue());
		}
		
		
		
	}

}
