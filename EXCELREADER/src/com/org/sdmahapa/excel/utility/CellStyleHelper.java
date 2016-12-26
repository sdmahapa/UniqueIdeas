package com.org.sdmahapa.excel.utility;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

public class CellStyleHelper {
	
	/**
	 * To set the font style
	 * @param font
	 * @param boldWeight
	 * @param color
	 * @param fontHeightInPoint
	 * @param fontName
	 * @param isItalic
	 * @param underLine
	 */
	public static void setFontStyle(Font font, short boldWeight, short color, short fontHeightInPoint, String fontName, boolean isItalic, byte underLine){
		if (font !=null){
			font.setBoldweight(boldWeight);
			font.setColor(color);
			font.setFontHeightInPoints(fontHeightInPoint);
			font.setFontName(fontName);
			font.setItalic(isItalic);
			font.setUnderline(underLine);
		}
	}
	
	/**
	 * To set the Cell Style
	 * @param styleHeading
	 * @param alignment
	 * @param verticalAlignment
	 * @param leftBoarder
	 * @param rightBoarder
	 * @param bottomBoarder
	 * @param topBoarder
	 * @param wrapText
	 */
	public static void setCellStyles(CellStyle styleHeading, short alignment, short verticalAlignment, short leftBoarder, short rightBoarder, short bottomBoarder, short topBoarder, boolean wrapText){
		if (styleHeading!=null){
			styleHeading.setAlignment(alignment);
			styleHeading.setVerticalAlignment(verticalAlignment);
			styleHeading.setBorderLeft(leftBoarder);
			styleHeading.setBorderRight(rightBoarder);
			styleHeading.setBorderBottom(bottomBoarder);
			styleHeading.setBorderTop(topBoarder);
			styleHeading.setWrapText(wrapText);
		}
		
	}
	
}
