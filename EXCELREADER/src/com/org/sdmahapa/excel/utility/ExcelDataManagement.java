package com.org.sdmahapa.excel.utility;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import com.org.sdmahapa.excel.entity.ExcelData;

/**
 * To manage the data into Excel file
 * @author SDMAHAPA
 *
 */
public class ExcelDataManagement {
	
	/**
	 * To fill the data into Excel
	 * @param workbook
	 * @param sheet
	 * @param listED
	 */
	public static void fillDataInToExcel(HSSFWorkbook workbook, HSSFSheet sheet, List<ExcelData> listED){
		
		int rownum=1;
		if (null!=workbook && null!=sheet && null!=listED){
			for (ExcelData ed : listED)
			{
				Row row = sheet.createRow(rownum++);
				
				Object [] objArr = {ed.getName(), ed.getDescription(), ed.getPrice(), ed.getDate()};
	            int cellnum = 0;
	            for (Object obj : objArr)
	            {
					
					Cell cell = row.createCell(cellnum++);
					CellStyle styleHeading = workbook.createCellStyle();
					if (obj instanceof String){
						CellStyleHelper.setCellStyles(styleHeading, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE, true);
						cell.setCellStyle(styleHeading);
						cell.setCellValue((String)obj);
					}
					else if (obj instanceof Double){
						CellStyleHelper.setCellStyles(styleHeading, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE, false);
						String format = String.valueOf((Double)obj*1.0f);
						cell.setCellStyle(styleHeading);
						cell.setCellValue(format);
					}
					else if (obj instanceof Date){
						CellStyleHelper.setCellStyles(styleHeading, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE, true);
						SimpleDateFormat sdf = new SimpleDateFormat("dd/mm/yyyy hh:MM:ssss");
						String format = sdf.format((Date)obj);
						cell.setCellStyle(styleHeading);
						cell.setCellValue(format);
					}
					else
						throw new IllegalArgumentException("Type Cast Exception!!!! It's only support String, Double");
	            }
			}
		}
		else
			throw new NullPointerException();
	}
	
}
