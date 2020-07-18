package com.hb.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FormatText {

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

	public FormatText(String FilePath) {
		try {
			file = new File(FilePath);
			workbook = new XSSFWorkbook(file);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}

	public void setFormatAsText() {
		//workbook = new XSSFWorkbook();

		sheet = workbook.getSheetAt(0);

		row = sheet.createRow(0);
		cell = row.createCell(0);
		format = workbook.createDataFormat();
		style = workbook.createCellStyle();
		style.setDataFormat(format.getFormat("Text"));
		cell.setCellStyle(style);
		oldStyle = cell.getCellStyle();
		newStyle = workbook.createCellStyle();
		newStyle.cloneStyleFrom(oldStyle);

	}

	public void runLoop(int rowNum, int colNum) {
		for (int i = 0; i < rowNum; i++) {
			row = sheet.createRow(i);

			for (int j = 0; j < colNum; j++) {
				cell = row.createCell(j);
				cell.setCellStyle(newStyle);
			}
		}

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
