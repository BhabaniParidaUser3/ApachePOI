package excelOperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingDataToFormulaCell1 {

	public static void main(String[] args) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Sheet1");
		XSSFRow row = sheet.createRow(0);
		row.createCell(0).setCellValue(100);
		row.createCell(1).setCellValue(100);
		row.createCell(2).setCellValue(100);
		row.createCell(3).setCellFormula("A1+B1+C1");

		String outputfilepath = ".\\dataFiles\\calc.xlsx";
		FileOutputStream outputStream = new FileOutputStream(outputfilepath);
		workbook.write(outputStream);
		outputStream.close();
		System.out.println("calc.xlsx created with formula cell");

	}

}
