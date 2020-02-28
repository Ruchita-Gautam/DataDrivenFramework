package util;

import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	// INTERVIEW QUESTION W/ ANSWER:
	// How do you get data from excel?
	// We have a class called ExcelReader. This class helps us to read the data from
	// excel file.
	// This class has three sections.
	// First is the set of global variables, second is the constructor
	// and third section consists of all the methods to get the cell data.
	// The global variables assign null values to certain classes that we need like
	// FileInputStream, XSSFWorkBook,
	// XSSFSheet, XSSFRow and XSSFCell.
	// The constructor instantiates all these classes and sets the values.
	// The methods will go in, find the cell values and return all these values.

	// Global Variables

	public String path;
	public FileInputStream fis = null;
	private XSSFWorkbook workbook = null;
	private XSSFSheet sheet = null;
	private XSSFRow row = null;
	private XSSFCell cell = null;

	// Constructor to initialize variables
	public ExcelReader(String path) {
		this.path = path;
		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			fis.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// Method to call the value
	public String getCellData(String sheetName, String colName, int rowNum) {
		// For Sheet
		int index = workbook.getSheetIndex(sheetName);
		int col_Num = 0;
		sheet = workbook.getSheetAt(index);

		// For Row
		row = sheet.getRow(0);
		for (int i = 0; i < row.getLastCellNum(); i++) {
			if (row.getCell(i).getStringCellValue().trim().equals(colName.trim())) {
				col_Num = i;
			}
		}

		// For Column
		row = sheet.getRow(rowNum - 1);
		cell = row.getCell(col_Num);
		return cell.getStringCellValue();
	}
}