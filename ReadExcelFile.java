package jatTASK8;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ReadExcelFile {

	public static void main(String[] args) {


		File file = new File("C:\\Users\\Anbarasan_SAM\\OneDrive\\Desktop\\WriteExcelSheet.xlsx");

		try (FileInputStream fis = new FileInputStream(file)) {
			Workbook workbook = new XSSFWorkbook(fis);

			Sheet sheet = workbook.getSheetAt(0);
			for (Row row : sheet) {
				for (Cell cell : row) {
					switch (cell.getCellType()) {
					case STRING:
						System.out.print(cell.getStringCellValue() + "\t");
						break;
					case NUMERIC:
						System.out.print(cell.getNumericCellValue() + "\t");
						break;
					default:
						System.out.print("Unknown type\t");
					}
				}
				System.out.println(); 
			}
			workbook.close();

		} catch (IOException e) {
			System.out.println("Error reading the Excel file: " + e.getMessage());
		}
	}
}
