package com.org.sdmahapa.excel.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileManagementHelper {

	private static File file = new File("Output\\Report.xls");
	
	/**
	 * To Write Data into Excel File.
	 * @param workbook
	 */
	public static void WriteToTheExcel(HSSFWorkbook workbook){
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(file);
			try {
				workbook.write(out);
			} catch (IOException e) {
				e.printStackTrace();
			}
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		finally{
			try {
				out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		
	}
	
	public static HSSFWorkbook ReadFromExcel() throws IOException{
		FileInputStream in = null;
		try{
			in = new FileInputStream(file);
			
			return new HSSFWorkbook(in);
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		return null;
	}
}
