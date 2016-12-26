package com.org.sdmahapa.excel.utility;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;

/*
 * Create a Header and manage
 */
public class ExcelHeaderHelper {

	/**
	 * To set the Header for Excel
	 * @param workbook
	 * @param sheet
	 * @param headerNames
	 */
	public static void setTheHeader(HSSFWorkbook workbook, HSSFSheet sheet,
			String... headerNames) {
		if (workbook != null && sheet != null && headerNames != null) {
			List<String> str = new ArrayList<String>();
			for (String arr : headerNames) {
				str.add(arr);
			}

			Row row = sheet.createRow(0);
			int cellnum = 0;
			for (String str1 : str) {
				Cell cell = row.createCell(cellnum++);
				CellStyle styleHeading = workbook.createCellStyle();
				Font font = workbook.createFont();
				CellStyleHelper.setFontStyle(font, (short) 8, (short) 0,
						(short) 12, HSSFFont.FONT_ARIAL, false, (byte) 0);
				styleHeading.setFont(font);
				CellStyleHelper.setCellStyles(styleHeading,
						CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER,
						CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE,
						CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE, true);
				cell.setCellStyle(styleHeading);
				cell.setCellValue(str1);
			}
		} else
			throw new NullPointerException();
	}
}
