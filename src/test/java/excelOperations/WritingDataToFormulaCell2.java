package excelOperations;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingDataToFormulaCell2 {

	public static void main(String[] args) throws IOException {

		String path=".\\dataFiles\\writeformulacell.xlsx";
		FileInputStream inputstream=new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(inputstream);
		XSSFSheet sheet = workbook.getSheetAt(0);
		sheet.getRow(7).getCell(2).setCellFormula("SUM(C2:C6)");	
		inputstream.close();
		FileOutputStream outputStream = new FileOutputStream(path);
		workbook.write(outputStream);
		outputStream.close();
		System.out.println("Done");

	}

}
