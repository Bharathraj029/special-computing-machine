package com.excel.interaction;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.master.interaction.ExcelReadData;

public class ReadDataFromExcel implements ExcelReadData {

	String colName;
	int rowNum;
	String excelPath;
	String sheeNme;

	public ReadDataFromExcel(String path, String sheet, int row, String column) {

		this.colName = column;
		this.rowNum = row;
		this.excelPath = path;
		this.sheeNme = sheet;

	}

	public String readData() {
		Logger log = Logger.getLogger(ReadDataFromExcel.class);
		FileInputStream fi = null;
		XSSFWorkbook wb=null;
		XSSFSheet sheet;
		XSSFRow row;
		XSSFCell cell;
		int colNum=0;
		try {
			fi = new FileInputStream(new File(excelPath));
			wb = new XSSFWorkbook(fi);
			sheet=wb.getSheet(sheeNme);
			row=sheet.getRow(0);
			for(int i=0;i<row.getLastCellNum();i++) {
				
				if(row.getCell(i).getStringCellValue().trim().equalsIgnoreCase(colName)) {
					colNum=i;
					break;
				}
			}
			
			row=sheet.getRow(rowNum);
			cell=row.getCell(colNum);
				
					switch(cell.getCellType()) {
					
					case Cell.CELL_TYPE_STRING:
					return cell.getStringCellValue();
					
					case Cell.CELL_TYPE_NUMERIC:
						return String.valueOf(cell.getNumericCellValue());
						
					case Cell.CELL_TYPE_BLANK:
						return String.valueOf(' ');
					
						default:
							
					}
					
					
				
			
			

		} catch (FileNotFoundException fn) {

			log.error(fn);
		} catch (IOException io) {
			log.error(io);
		}
		catch(Exception e1) {
			log.error(e1);
		}

		finally {

			try {

				fi.close();

			} catch (IOException e) {

				log.error(e);
			}
		}

		return null;
	}

}
