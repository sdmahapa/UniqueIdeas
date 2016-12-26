package com.org.sdmahapa.excel;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import com.org.sdmahapa.excel.entity.ExcelData;
import com.org.sdmahapa.excel.utility.ExcelDataManagement;
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
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		
		HSSFSheet sheet = workbook.createSheet("Expanse_Data");
		
		//data to be filled.
		
		//Only for header
		try
		{
			ExcelHeaderHelper.setTheHeader(workbook, sheet, "Name", "Description", "Amount", "TimeStamp");
			
			//data
			List<ExcelData> led = new ArrayList<ExcelData>();
			ExcelData ed = new ExcelData();
			ed.setName("Bike EMI");
			ed.setDescription("Royal Enfield bike, every month needs to pay the load amount");
			ed.setPrice(5000.00);
			ed.setDate(new Date());
			led.add(ed);
			
			ExcelDataManagement.fillDataInToExcel(workbook, sheet, led);
		
		
		
            
            //for (int i=0; i<cellnum; i++)
            	//sheet.autoSizeColumn(cellnum);
			
		
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
