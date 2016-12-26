package com.org.sdmahapa.excel.utility;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class FileManagementHelper {

	private static File file = new File("Output\\Report.xls");
	private static FileOutputStream out = null;
	
	public static void WriteToTheExcel(HSSFWorkbook workbook){
		
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
}
