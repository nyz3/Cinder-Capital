package excelreadwrite;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
	
	public static void main(String[] args) throws IOException {
		
		//inputstream "fileIn" is the file we are reading data/info from
		FileInputStream fileIn = new FileInputStream(new File("FinancialData.xlsx"));
		//outputstream "fileOut" is the file we are writing out manipulated data to
		FileOutputStream fileOut = new FileOutputStream(new File("FinancialDataAnalysis.xlsx"));
		
		//create workbook from the given excel file to allow access to cells/data, e.g getNumberofSheets()
		XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
		
		//gets first sheet of the excel file, edit later for numerous sheets
		XSSFSheet sheet = workbook.getSheetAt(0); 
		
		
		workbook.close();
		fileOut.close();
		fileIn.close();
		
		
		
		
		
	}
}
