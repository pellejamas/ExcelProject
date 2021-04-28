package com.demo.ExcelProject;

import java.io.FileOutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Createinvoice {

	public static void main(String[] args) {
		try {
			//Create workbook in .xls format
			Workbook workbook = new HSSFWorkbook();
			//For .xslx workbooks use XSSFWoorkbok();
			//Create Sheet
			Sheet sh = workbook.createSheet("Invoices");
			//Create top row with column headings
			String[] columnHeadings = {"Item Id", "Item Name", "Qty", "Item Price", "Sold Date"};
			//We want to make it bold with a foreground color
			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setFontHeightInPoints((short)12);
			headerFont.setColor(IndexedColors.BLACK.index);
			//create a CellStyle with the font
			CellStyle headerStyle = workbook.createCellStyle();
			headerStyle.setFont(headerFont);
			headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
			//Create the header row
			Row headerRow = sh.createRow(0);
			//Iterate over the column to Create columns
			for(int i=0;i<columnHeadings.length;i++){
				Cell cell = headerRow.createCell(i);
				cell.setCellValue(columnHeadings[i]);
				cell.setCellStyle(headerStyle);
			}
			//Fill data
			ArrayList<Invoices> a = createData();
			CreationHelper creationHelper = workbook.getCreationHelper();
			CellStyle dateStyle = workbook.createCellStyle();
			dateStyle.setDataFormat(creationHelper.createDataFormat().getFormat("MM/dd/yyyy"));
			int rownum =1;
			for(Invoices i : a) {
				Row row = sh.createRow(rownum++);
				row.createCell(0).setCellValue(i.getItemId());
				row.createCell(1).setCellValue(i.getItemName());
				row.createCell(2).setCellValue(i.getItemQty());
				row.createCell(3).setCellValue(i.getTotalPrice());
				Cell dateCell = row.createCell(4);
				dateCell.setCellValue(i.getItemSoldDate());
				dateCell.setCellStyle(dateStyle);
			}
			//Autosize columns
			for(int i=0;i<columnHeadings.length;i++) {
				sh.autoSizeColumn(i);
			}
			@SuppressWarnings("unused")
			Sheet sh2 = workbook.createSheet("Second");
			//Write the output to file
			FileOutputStream fileOut = new FileOutputStream("C:\\Users\\snick\\Desktop\\excel4.xls");
			workbook.write(fileOut);
			fileOut.close();
			workbook.close();
			System.out.println("Completed");
		}
		catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static ArrayList<Invoices> createData() throws ParseException {ArrayList<Invoices> a = new ArrayList<Invoices>(); 
		a.add(new Invoices(1, "Book", 2, 10.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(2, "Table", 1, 50.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(3, "Lamp", 5, 100.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(4, "Pen", 100, 20.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(5, "Book", 2, 10.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(6, "Table", 1, 50.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(7, "Lamp", 5, 100.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(8, "Pen", 100, 20.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(9, "Book", 2, 10.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(10, "Table", 1, 50.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(11, "Lamp", 5, 100.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(12, "Pen", 100, 20.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(13, "Book", 2, 10.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(14, "Table", 1, 50.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		a.add(new Invoices(15, "Lamp", 5, 100.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
		return a;
	}
}