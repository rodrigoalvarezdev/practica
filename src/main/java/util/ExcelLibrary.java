package util;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelLibrary {
	public static void readAllExcel() throws IOException {
		try {
			FileInputStream f = new FileInputStream("data\\TestData.xlsx");
			XSSFWorkbook libro = new XSSFWorkbook(f); 
			
			XSSFSheet hoja = libro.getSheet("credentials");
			System.out.println(hoja.getLastRowNum());
			Iterator<Row> filas = hoja.rowIterator();
			Iterator<Cell> celdas;
			Row fila;
			Cell celda;
			while(filas.hasNext()) {
				fila = filas.next();
				celdas = fila.cellIterator();
				
				while(celdas.hasNext()) {
					celda = celdas.next();
					System.out.println(celda.getStringCellValue());
				}
			}
			libro.close();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
	
	public static void readCell() throws IOException {
		try {
			FileInputStream f = new FileInputStream("data\\TestData.xlsx");
			XSSFWorkbook libro = new XSSFWorkbook(f);
			XSSFSheet hoja = libro.getSheet("credentials");
			XSSFRow fila = hoja.getRow(1);
			XSSFCell celda = fila.getCell(0);
			System.out.println(celda.getStringCellValue());
			libro.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public static void readColumn() throws IOException {
		try {
			FileInputStream f = new FileInputStream("data\\TestData.xlsx");
			XSSFWorkbook libro = new XSSFWorkbook(f); 
			
			XSSFSheet hoja = libro.getSheet("credentials");
			
			Iterator<Row> filas = hoja.rowIterator();
			Row fila;
			while(filas.hasNext()) {
				fila = filas.next();
				Cell celda = fila.getCell(0);
				String celValue = celda.getStringCellValue();
				System.out.println(celValue);
			}
			libro.close();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
	
	public static void newColumn() throws IOException {
		
		try {
			FileInputStream inputStream = new FileInputStream("data\\TestData.xlsx");
			
			XSSFWorkbook newWorkbook = new XSSFWorkbook(inputStream);
			
			XSSFSheet newSheet = newWorkbook.getSheet("credentials");

			
			int rowCount = newSheet.getLastRowNum() - newSheet.getFirstRowNum();
			System.out.println(rowCount);
			
			int rowTotal = newSheet.getLastRowNum();
			System.out.println(rowTotal);
			
			for (int i = 0; i <= 2; i++) {
				XSSFRow newRow = newSheet.getRow(i);
				String newCell = newRow.getCell(0).getStringCellValue();
				System.out.println(newCell);
			}
			newWorkbook.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
