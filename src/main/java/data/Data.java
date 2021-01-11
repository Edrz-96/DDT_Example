package data;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Data {
	private static XSSFSheet sheet;

	private static XSSFWorkbook book;

	private static XSSFCell cell;

	private static XSSFRow row;

	private static DataFormatter dataFormatter;

	// Provides sheet names

	public static List<String> getSheetNames() {
		List<String> list = new ArrayList<>();
		for (int x = 0; x <= book.getNumberOfSheets() - 1; x++) {
			list.add(book.getSheetName(x));
		}
		return list;

	}

	// Provides cell quantity, included for syntax

	public static int getCellQuanity() {
		int no = 0;
		for (int i = 0; i <= sheet.getPhysicalNumberOfRows() - 1; i++) {
			for (int j = 0; j <= sheet.getRow(i).getPhysicalNumberOfCells() - 1; j++)
				no++;
		}
		return no;
	}

	// Provides row quantity, included for syntax

	public static int getRowQuanity() {
		return sheet.getPhysicalNumberOfRows();

	}

	// Provides cell data when passed row and col index

	public static String getSingleCellData(int row, int col) throws Exception {

		String data = sheet.getRow(row).getCell(col).getStringCellValue();

		return data;
	}

	// Provides cell data when passed row and col index and sheet

	public static String getSingleCellData(int row, int col, String sheetName) throws Exception {
		sheet = book.getSheet(sheetName);
		String data = sheet.getRow(row).getCell(col).getStringCellValue();

		return data;
	}

	// Prints all cell data to console, useful in smaller tests

	public static void getAllCellData(String sheetName) throws Exception {

		sheet = book.getSheet(sheetName);

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
	}

	// Updates a cell in the database provided

	public static void setSingleCellData(String input, int rowNo, int colNo, String path, String file)
			throws Exception {

		book.setMissingCellPolicy(MissingCellPolicy.RETURN_BLANK_AS_NULL);

		row = sheet.getRow(rowNo);

		cell = row.getCell(colNo);

//		row.createCell(colNo).setCellValue(input);

		if (cell == null) {
			row.createCell(colNo);
			cell.setCellValue(input);
		} else {
			cell.setCellValue(input);
		}

		FileOutputStream fileOut = new FileOutputStream(path + file);
		book.write(fileOut);
		fileOut.flush();
		fileOut.close();

	}

	// Opens data at input with sheet

	public static void setExcelFile(String file, String path, String sheetName) throws Exception {

		FileInputStream fileIn = new FileInputStream(path + file);

		book = new XSSFWorkbook(fileIn);

		sheet = book.getSheet(sheetName);
	}

	// Opens data at input
	public static void setExcelFile(String file, String path) throws Exception {

		FileInputStream fileIn = new FileInputStream(path + file);

		book = new XSSFWorkbook(fileIn);

	}

	// Selection method for sheet
	public static XSSFSheet setSheet(String sheetName) {
		return sheet = book.getSheet(sheetName);

	}
}
