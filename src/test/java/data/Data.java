package data;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class Data {
	XSSFWorkbook workbook;
	DataFormatter dataFormatter = new DataFormatter();

	@Test
	public void loginTest() throws Exception {

		// Reading
		FileInputStream file = new FileInputStream("src/test/resources/data/test.xlsx");
		workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		XSSFSheet sheetTwo = workbook.getSheetAt(1);

		System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets:");

		for (int x = 0; x <= workbook.getNumberOfSheets() - 1; x++) {
			System.out.println(workbook.getSheetName(x));
		}
		
		
		String name = dataFormatter.formatCellValue(sheet.getRow(1).getCell(0));
		System.out.println(name);
		
		
		Iterator<Row> rowIterator = sheet.rowIterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();

			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				String cellValue = dataFormatter.formatCellValue(cell);
				System.out.format("%32s%16s", "", cellValue);
			}
			System.out.println();
		}

		// Update Sheet names
		workbook.setSheetName(0, "Login Credentials");
		workbook.setSheetName(1, "Assert Data");

		// Create headers for sheet 1
		Row header = sheet.createRow(0);
		header.createCell(0).setCellValue("Username");
		header.createCell(1).setCellValue("Password");

		// Create headers for sheet 2
		Row headers = sheetTwo.createRow(0);
		headers.createCell(0).setCellValue("Assertion");
		headers.createCell(1).setCellValue("Status");

		// Writing
		FileOutputStream fileOut = new FileOutputStream("src/test/resources/data/test.xlsx");
		workbook.write(fileOut);
		fileOut.flush();
		fileOut.close();
		workbook.close();
		file.close();

	}

}
