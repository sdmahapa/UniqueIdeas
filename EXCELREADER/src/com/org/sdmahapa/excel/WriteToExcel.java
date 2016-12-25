package com.org.sdmahapa.excel;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;


/**
 * 
 * @author SDMAHAPA
 *
 */
public class WriteToExcel{
	public static void main(String args[])
	{
		@SuppressWarnings("resource")
		HSSFWorkbook workbook = new HSSFWorkbook();
		
		HSSFSheet sheet = workbook.createSheet("Expanse_Data");
		
		//data to be filled.
		
		Map<String, Object[]> data = new HashMap<String, Object[]>();
		data.put("1", new Object[] {"Name", "Description", "Amount"});
		data.put("2", new Object[] {"Flat_Rent", "KMS Residency Rent", "6500"});
		data.put("3", new Object[] {"Cook", "Cooking aunty", "1200"});
		data.put("4", new Object[] {"Bike_EMI", "Tata Capital for loan", "5000"});
		data.put("5", new Object[] {"Mobile Bill", "VodafonePostpaid", "1200"});
		
		Set<String> keyset = data.keySet();
		int rownum = 0;
		
		for (String key : keyset)
		{
			Row row = sheet.createRow(rownum++);
			Object[] objarr = data.get(key);
			int cellnum = 0;
			
			for(Object obj : objarr)
			{
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof String)
					cell.setCellValue((String)obj);
				else if (obj instanceof Integer)
					cell.setCellValue((Integer)obj);
				else
					throw new IllegalArgumentException("Type Cast Exception!!!! It's only support String and Integer");
			}
		}
		try
		{
			FileOutputStream out = new FileOutputStream(new File("C:\\Users\\SDMAHAPA\\workspace\\EXCELREADER\\Expanse_sheet.xls"));
			workbook.write(out);
			out.close();
			System.out.println("Expanse_sheet.xlsx written successfully on disk!");
		}
		catch (IOException e)
		{
			e.printStackTrace();
			System.out.println("IOException::"+e);
		}
		catch(Exception e)
		{
			e.printStackTrace();
			System.out.println("General Exception::"+e);
		}
	}

}
