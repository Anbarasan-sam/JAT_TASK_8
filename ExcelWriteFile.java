package jatTASK8;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriteFile {

	public static void main(String[] args) {
		Workbook workbook = new XSSFWorkbook();  
		Sheet sheet = workbook.createSheet("Sheet1");
		Row headerRow = sheet.createRow(0);
		Cell cell1 = headerRow.createCell(0);
		cell1.setCellValue("Name");

		Cell cell2 = headerRow.createCell(1);
		cell2.setCellValue("Age");

		Cell cell3 = headerRow.createCell(2);
		cell3.setCellValue("Email");

		Object[][] data = {
				{"John Doe", 30, "john@test.com"},
				{"Jane Doe", 28, "john@test.com"},
				{"Bob Smith", 35, "jacky@example.com"},
				{"Swapnil", 37, "swapnil@example.com"}
		};

		int rowNum = 1;
		for (Object[] rowData : data) {
			Row row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue((String) rowData[0]); 
			row.createCell(1).setCellValue((Integer) rowData[1]);
			row.createCell(2).setCellValue((String) rowData[2]); 
		}

		File file = new File("WriteExcelSheet.xlsx");

		try (FileOutputStream fileOut = new FileOutputStream(file)) {
			workbook.write(fileOut);
			System.out.println("Excel file created successfully!");
		} catch (IOException e) {
			System.out.println("Error writing Excel file: " + e.getMessage());
		} finally {
			try {
				workbook.close();  
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}
