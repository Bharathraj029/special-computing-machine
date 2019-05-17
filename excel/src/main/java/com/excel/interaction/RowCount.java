package com.excel.interaction;

import java.io.File;
import java.io.FileInputStream;
import org.apache.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.master.interaction.ExcelRowCount;



public class RowCount implements ExcelRowCount {

	String filePath;
	String sheetName;

	public RowCount(String path, String sheetname) {

		this.filePath = path;
		this.sheetName = sheetname;
	}

	public int rowFetched() {
		Logger logger = Logger.getLogger(RowCount.class.getName());
		int count = 0;
		try {
			FileInputStream fi = new FileInputStream(new File(filePath));

			XSSFWorkbook wb = new XSSFWorkbook(fi);
			XSSFSheet sheet = wb.getSheet(sheetName);
			if (sheet.getSheetName().isEmpty()) {

				logger.error("Sheet name is not available");

			} else {

				if (sheet.getLastRowNum() > 0) {

					count = sheet.getLastRowNum();

				} else {

					logger.error("Row doesn't exist");
				}
			}

		} catch (Exception e) {
			logger.error(e);
		}
		return count;

	}

}
