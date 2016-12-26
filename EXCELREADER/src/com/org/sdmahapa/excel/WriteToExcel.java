package com.org.sdmahapa.excel;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;

import com.org.sdmahapa.excel.entity.ExcelData;
import com.org.sdmahapa.excel.utility.CellStyleHelper;
import com.org.sdmahapa.excel.utility.ExcelHeaderHelper;
import com.org.sdmahapa.excel.utility.FileManagementHelper;


/**
 * 
 * @author SDMAHAPA
 *
 */
public class WriteToExcel{
	public static void main(String args[])
	{
		Set<String> keyset = null;
		int cellnum = 0;
		int rownum = 1;
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		
		HSSFSheet sheet = workbook.createSheet("Expanse_Data");
		
		//data to be filled.
		
		//Only for header
		ExcelHeaderHelper.setTheHeader(workbook, sheet, "Name", "Description", "Amount", "TimeStamp");
		
		//data
		Map<String, ExcelData> data = new HashMap<String, ExcelData>();
		ExcelData ed = new ExcelData();
		ed.setName("Bike EMI");
		ed.setDescription("Royal Enfield bike, every month needs to pay the load amount");
		ed.setPrice(5000.00);
		ed.setDate(new Date());
		data.put("2", ed);
		
		
		keyset = data.keySet();
		
		
		for (String key : keyset)
		{
			Row row = sheet.createRow(rownum++);
			ExcelData dataArr = data.get(key);
			
			Object [] objArr = {dataArr.getName(), dataArr.getDescription(), dataArr.getPrice(), dataArr.getDate().toString()};
            cellnum = 0;
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
					cell.setCellStyle(styleHeading);
					cell.setCellValue((Double)obj);
				}
				else
					throw new IllegalArgumentException("Type Cast Exception!!!! It's only support String, Double");
            }
            
            //for (int i=0; i<cellnum; i++)
            	//sheet.autoSizeColumn(cellnum);
			
		}
		try
		{
			FileManagementHelper.WriteToTheExcel(workbook);
			System.out.println("Report.xlsx written successfully on disk! Please find in following path");
		}
		catch(Exception e)
		{
			e.printStackTrace();
			System.out.println("General Exception::"+e);
		}
	}

}
