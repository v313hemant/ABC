package ExcelConversion;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/* * * * * * * * * * * * * * * * * * * * * * * * * */
/*  	Developer : Richa Verma					   */
/*  	Date : 06/01/2017 						   */
/* * * * * * * * * * * * * * * * * * * * * * * * * */

public class CreateStyles {
	
	public static CellStyle styleHeader = null;
	public static CellStyle styleBanner = null;
	public static CellStyle greyHeader = null;
	public static CellStyle styleBordered = null;
	public static CellStyle greyFill = null;
	
	public static CellStyle dateStyle = null;
	public static CellStyle date2Style = null;
	public static CellStyle timestampStyle = null;
	public static CellStyle currencyStyle = null;
	public static CellStyle Style_number_custom_1 = null;
	public static CellStyle timeStyle = null;
	public static DataFormat currencyFormat = null;
	public static DataFormat Format_Number_Custom_1 = null;
	
	public static void setBorders(CellStyle borderStyle) {
		
		borderStyle.setBorderBottom(CellStyle.BORDER_THIN);
		borderStyle.setBorderTop(CellStyle.BORDER_THIN);
		borderStyle.setBorderRight(CellStyle.BORDER_THIN);
		borderStyle.setBorderLeft(CellStyle.BORDER_THIN);
	}
	
	public static void createStyles(String fileFormat, HSSFWorkbook wb, SXSSFWorkbook wbx, int sheetCount, String reportName) {

		Font fontHeader = null;
		Font fontBanner = null;
		Font fontGreyHeader = null;
		
		try {
			if (fileFormat.toUpperCase().equals("XLSX")) {
				
				CreationHelper createHelper = wbx.getCreationHelper();
				
				styleHeader = wbx.createCellStyle();
		        styleBanner = wbx.createCellStyle();
		        greyHeader = wbx.createCellStyle();
		        styleBordered = wbx.createCellStyle();
		        greyFill = wbx.createCellStyle();
		        
		        fontHeader = wbx.createFont();
		        fontBanner = wbx.createFont();
		        fontGreyHeader = wbx.createFont();
		        
		        currencyStyle = wbx.createCellStyle();
				currencyFormat = wbx.createDataFormat();
				currencyStyle.setDataFormat(currencyFormat.getFormat("£#,##0.00;[Red]-£#,##0.0000"));
				
				Style_number_custom_1 = wbx.createCellStyle();
				Format_Number_Custom_1 = wbx.createDataFormat();
				Style_number_custom_1.setDataFormat(Format_Number_Custom_1.getFormat("#,##0.00;[Red](#,##0.00);-"));
	
				dateStyle = wbx.createCellStyle();
				dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/MM/yyyy"));
	
				date2Style = wbx.createCellStyle();
				date2Style.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-MM-dd"));
				
				timestampStyle = wbx.createCellStyle();
				timestampStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/MM/yyyy hh:mm"));
				
				timeStyle = wbx.createCellStyle();
				timeStyle.setDataFormat(createHelper.createDataFormat().getFormat("HH:mm:ss"));
							
			} else if (fileFormat.toUpperCase().equals("XLS")) {
				
				CreationHelper createHelper = wb.getCreationHelper();
				
				styleHeader = wb.createCellStyle();
		        styleBanner = wb.createCellStyle();
		        greyHeader = wb.createCellStyle();
		        styleBordered = wb.createCellStyle();
		        greyFill = wb.createCellStyle();
		        
		        fontHeader = wb.createFont();
		        fontBanner = wb.createFont();
		        fontGreyHeader = wb.createFont();
				
		        currencyStyle = wb.createCellStyle();
				currencyFormat = wb.createDataFormat();
				currencyStyle.setDataFormat(currencyFormat.getFormat("£#,##0.00;[Red]-£#,##0.0000"));
				
				Style_number_custom_1 = wb.createCellStyle();
				Format_Number_Custom_1 = wb.createDataFormat();
				Style_number_custom_1.setDataFormat(Format_Number_Custom_1.getFormat("#,##0.00;[Red](#,##0.00);-"));
				
				dateStyle = wb.createCellStyle();
				dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/MM/yyyy"));
				
				date2Style = wb.createCellStyle();
				date2Style.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-MM-dd"));
				
				timestampStyle = wb.createCellStyle();
				timestampStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/MM/yyyy hh:mm:ss"));
				
				timeStyle = wb.createCellStyle();
				timeStyle.setDataFormat(createHelper.createDataFormat().getFormat("HH:mm:ss"));
				
			} else {
				System.out.println("Invalid File Format");
			}
	
	        // Creating style variable for headers
			fontHeader.setBoldweight(Font.BOLDWEIGHT_BOLD);
	        styleHeader.setFont(fontHeader);
	        styleHeader.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
	    	styleHeader.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    	
	    	fontGreyHeader.setBoldweight(Font.BOLDWEIGHT_BOLD);
	    	greyHeader.setFont(fontGreyHeader);
	    	greyHeader.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
	        greyHeader.setFillPattern(CellStyle.SOLID_FOREGROUND);
	        
	        greyFill.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
	        greyFill.setFillPattern(CellStyle.SOLID_FOREGROUND);
	        
	        // Creating style variable for banner
	        fontBanner.setBoldweight(Font.BOLDWEIGHT_BOLD);
	        styleBanner.setFont(fontBanner);
	       
	        // Special handling for STD-AF - applying cell borders
			if (reportName.equals("STD-AF") && sheetCount == 1) {
				CreateStyles.setBorders(styleBanner);
				CreateStyles.setBorders(greyHeader);
				CreateStyles.setBorders(currencyStyle);
				CreateStyles.setBorders(dateStyle);
				CreateStyles.setBorders(date2Style);
				CreateStyles.setBorders(timestampStyle);
				CreateStyles.setBorders(timeStyle);
				CreateStyles.setBorders(styleHeader);
				CreateStyles.setBorders(styleBordered);
				CreateStyles.setBorders(greyFill);
			}
		} catch (Exception e) {
			System.err.println("Exception - " + e.getMessage());
			System.exit(1);
		}
	}
}
