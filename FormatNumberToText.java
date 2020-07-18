package com.hb.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FormatNumberToText {
	int Counter = 0;
	File file = null;
	XSSFWorkbook workbook = null;
	XSSFSheet sheet = null;
	XSSFRow row = null;
	XSSFCell cell = null;
	XSSFDataFormat format = null;
	XSSFCellStyle style = null;
	XSSFCellStyle oldStyle = null;
	XSSFCellStyle newStyle = null;
	FileOutputStream fos = null;
	FileInputStream fis = null;
	Object obj = null;

	public FormatNumberToText(String FilePath) {
		try {
			file = new File(FilePath);
			workbook = new XSSFWorkbook(file);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}

	public void ConvertNumberToText(int colNum) {
		sheet = workbook.getSheetAt(0);
		int numberOfRows = sheet.getLastRowNum();
		System.out.println(numberOfRows);
		for (int rowNum = 1; rowNum <= numberOfRows; rowNum++) {
			row = sheet.getRow(rowNum);
			cell = row.getCell(colNum);
			if(cell!=null)
			{
			if (cell.getCellType() == CellType.NUMERIC && Counter == 0) {
				Counter++;
				obj = cell.getNumericCellValue();
				format = workbook.createDataFormat();
				style = workbook.createCellStyle();
				style.setDataFormat(format.getFormat("Text"));
				cell.setCellStyle(style);
				cell.setCellValue(new BigDecimal(obj.toString()).setScale(0, RoundingMode.UNNECESSARY).toPlainString());
				oldStyle = cell.getCellStyle();
				newStyle = workbook.createCellStyle();
				newStyle.cloneStyleFrom(oldStyle);
				
			}
			if (cell.getCellType() == CellType.NUMERIC && Counter != 0) {

				
				obj = cell.getNumericCellValue();
				cell.setCellStyle(newStyle);
				cell.setCellValue(new BigDecimal(obj.toString()).setScale(0, RoundingMode.UNNECESSARY).toPlainString());
				
			}
		}}

		try {
			fos = new FileOutputStream(file, true);
			workbook.write(fos);
			workbook.close();
			fos.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
