package ExcelConversion;

import ExcelConversion.CreateStyles;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.nio.ReadOnlyBufferException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import au.com.bytecode.opencsv.CSVReader;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/* * * * * * * * * * * * * * * * * * * * * * * * * */
/*  	Developer : Richa Verma					   */
/*  	Date : 28/12/2016 						   */
/* * * * * * * * * * * * * * * * * * * * * * * * * */

public class ConvertAndFormatExcel {
	
	CSVToEXCEL cte = new CSVToEXCEL();
	
//	static String headerPos=null;
//	static String Zip_password=null;
	static HSSFWorkbook wb = new HSSFWorkbook();
	static SXSSFWorkbook wbx = new SXSSFWorkbook();
	static int check_std_af=0;
	static FileOutputStream fileOut = null;
	private BufferedReader formatBR;
	
	public int temp = 0;
	
	// Add all handled data types here
	public static enum Datatypes {
		 STRING, NUMBER, DATE, DATE2, CURRENCY, TIMESTAMP, TIME, NUMBERC1
	}
			
	public void convertAndFormat (String Flist, String OutFname, String Ftype, String bannerLen, String headerPos) throws IOException {
		
//		//If password
//		if(headerPos_pass.contains("#")){
//			headerPos=headerPos_pass.substring(0,headerPos_pass.indexOf('#'));
//			Zip_password=headerPos_pass.substring(headerPos_pass.indexOf('#')+1).trim();
//		}
//		else{
//			headerPos=headerPos_pass;
//			Zip_password="NA";
//		}
		
		
		// Declaring variables
		String reportName = CSVToEXCEL.reportName.trim().toUpperCase();
		List<String> lines = new ArrayList<String>();
		List<String> cellFormatArray = new ArrayList<String>();
		
		Cell cell = null;
		Row row = null;
		Sheet sheet = null;
		CreationHelper helper = null;		
		ArrayList<Integer> maxNoChar = new ArrayList<Integer>();
		
		String InputString = null;
		String FileName = null;
		String infile = null;
		String WorksheetName = null;
		String OutputFileName = "";
		String ReportFormatFile = null;
		String[] parts = null;
		String[] line = null;
		
		String frmtFile = null;
		String[] splitFormat = null;
		String cellFormat = null;
		String cellFrmtValue = null;
		int headerPosition = Integer.parseInt(headerPos);
		int bannerLength = Integer.parseInt(bannerLen);
		int cellCount = 0;
		int sheetCount = 0;
		int rowCount = 0;
		int blockCheck = 0;
		
		CellStyle newStyle = null;
		Font newFont = null;
		
		try {
			FileReader fileReader = new FileReader(Flist);
			BufferedReader bufferedReader = new BufferedReader(fileReader);
			
			String FOutName = OutFname;
			String FileType = Ftype.toLowerCase();
			
			if (FileType.equals("xls")) {
				OutputFileName = FOutName + ".xls";
			} else if (FileType.equals("xlsx")) {
				OutputFileName = FOutName + ".xlsx";
				wbx.setCompressTempFiles(true);
			} else {
				System.out.println("Invalid Output File Type entered - " + FileType);
				System.exit(1);
			}
			
			// Read list of input CSVs in an array lines
			while ((infile = bufferedReader.readLine()) != null) {
				lines.add(infile);
			}
		
			System.out.println("Number of CSVs to be merged to " + FileType.toUpperCase() + " : " + lines.size());
			
			for (int i = 0; i < lines.size(); i++) {
				
				InputString = lines.get(i);
				parts = InputString.split("\\|");
				FileName = parts[0];
				WorksheetName = parts[1];
				ReportFormatFile = parts[2];
								
				FileReader formatFR = new FileReader(ReportFormatFile);
				formatBR = new BufferedReader(formatFR);
				
				// Reset the array and Read Format type of each cell in an array
				cellFormatArray.clear();
				while ((frmtFile = formatBR.readLine()) != null) {
					splitFormat = frmtFile.split("\\,");
					cellFormat = splitFormat[1];
					cellFormatArray.add(cellFormat);
				}
				
				System.out.println("Input CSV Name : " + FileName + "\nWorksheet Name : " + WorksheetName);
				
				if (FileType.equals("xls")) {
					helper = wb.getCreationHelper();
					sheet = wb.createSheet(WorksheetName); // Create XLS Worksheet
					sheetCount = sheetCount + 1;
					newStyle = wb.createCellStyle();
					newFont = wb.createFont();
				} else if (FileType.equals("xlsx")) {
					helper = wbx.getCreationHelper();
					sheet = wbx.createSheet(WorksheetName); // Create XLSX Worksheet
					sheetCount = sheetCount + 1;
					newStyle = wbx.createCellStyle();
					newFont = wbx.createFont();
				} else {
					System.out.print("Invalid File Type");
				}
							
				// Creating Formatting variables when FormatFlag is TRUE
				CreateStyles.createStyles(Ftype, wb, wbx, sheetCount, reportName);
				CSVReader reader = null;
				
				try {
					reader = new CSVReader(new FileReader(FileName)); // Reading Input CSV files
					long r = 0;
					int flag = 0;
					
					while ((line = reader.readNext()) != null) {
						
						row = sheet.createRow((int) r++);
						
						if ( reportName.equals("STD-AF") && sheetCount == 1) {
							check_std_af=1;
							if ( line[0].equals(null) || line[0].trim().equals("") ) { 
							
								rowCount = 0;
								blockCheck++;
							} else {
								rowCount++;
							}
							
							for (int j = 0; j < line.length; j++) {
								flag = 0;
								cell = row.createCell(j);
								cellFrmtValue = cellFormatArray.get(j).trim();
								
								newFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
								
								if (rowCount > 0) {
									
									if (blockCheck == 0) {
										cell = setCellValueAndStyle(cell, helper, flag, cellFrmtValue, line[j]);
										if ( j == 0 ) {
											cell.setCellStyle(CreateStyles.styleBanner);
										} else if ( j == 1 ) {
											cell.setCellStyle(CreateStyles.styleBordered);
										}
									} else if (blockCheck > 0) {
										if (rowCount == 1) {
											cell = setCellValueAndStyle(cell, helper, flag, cellFrmtValue, line[j]);
											if ( j == 0 ) {
												cell.setCellStyle(CreateStyles.greyHeader);
											}
										} else if (rowCount == 2) {
											cell = setCellValueAndStyle(cell, helper, flag, cellFrmtValue, line[j]);
											cell.setCellStyle(CreateStyles.styleHeader);
										} else if (line[0].toUpperCase().trim().equals("GRAND TOTAL")) {
											
											if (j == 0) {
												cell = setCellValueAndStyle(cell, helper, flag, cellFrmtValue, line[j]);
											} else {
												cell = setCellValueAndStyle(cell, helper, 1, cellFrmtValue, line[j]);
											}
											
											if (j == 7) {
												newStyle.cloneStyleFrom(cell.getCellStyle());
												newStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
												newStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
												cell.setCellStyle(newStyle);
											}
											
											if ( cellFrmtValue.equals("STRING") || cellFrmtValue.equals("NUMBER")) {
												cell.setCellStyle(CreateStyles.styleBordered);
												// newStyle.setFont(newFont);
											}
										} else {
											cell = setCellValueAndStyle(cell, helper, 1, cellFrmtValue, line[j]);
											
											if (j == 7) {
												newStyle.cloneStyleFrom(cell.getCellStyle());
												newStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
												newStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
												cell.setCellStyle(newStyle);
											} else if (cellFrmtValue.equals("STRING") || cellFrmtValue.equals("NUMBER")) {
												cell.setCellStyle(CreateStyles.styleBordered);
											}
										}  
									}
									
									// Count of cells in a Row for width adjustments
									cellCount = row.getPhysicalNumberOfCells();
									
									if (maxNoChar.size() < line.length) {
										maxNoChar.add(j, temp);
									} else if (temp > maxNoChar.get(j)) {
										maxNoChar.set(j, temp);
									}
								}
							}
						
						} else {
							check_std_af=2;
							if (r == headerPosition+1) { 
								flag = 1;  // Format report data as per the datatypes defined
							}
							
							for (int j = 0; j < line.length; j++) {
								
								if(line.length==1 && line[0].equals("")){
									
								}
								else{
									cell = row.createCell(j);
									cellFrmtValue = cellFormatArray.get(j).trim();
									
									
									// ----- Calling function to set the cell value and apply style ----- //
									cell = setCellValueAndStyle(cell, helper, flag, cellFrmtValue, line[j]);
									// ------------------------------------------------------------------ //
																	
									if (r > bannerLength) {
										// Reading maximum length for each row for column width setting
										if (maxNoChar.size() < line.length) {
											maxNoChar.add(j, temp);
										} else if (temp > maxNoChar.get(j)) {
											maxNoChar.set(j, temp);
										}
									}
									
									// Formatting Banner - displaying banner titles in bold
									if ((j == 0) && (r <= bannerLength)) {
										cell.setCellStyle(CreateStyles.styleBanner);
									}
									
									// Header Formatting (Bold and coloured background)
									if ((r == headerPosition)) {
										cell.setCellStyle(CreateStyles.styleHeader);
										cellCount = row.getPhysicalNumberOfCells();
									}	
									
									
								}
								
							
							}	
						}
					}
					
					if(check_std_af==2 && headerPosition > 0){
						// Header Formatting (Add AutoFilters) and Freeze Pane
						
//						System.out.println("A"+ headerPosition + ":"+ CellReference.convertNumToColString(cellCount-1)+headerPosition);
//						System.out.println(CellRangeAddress.valueOf("A"+ headerPosition + ":"+ CellReference.convertNumToColString(cellCount-1)+headerPosition));
						
						sheet.setAutoFilter(CellRangeAddress.valueOf("A"+ headerPosition + ":"+ CellReference.convertNumToColString(cellCount-1)+headerPosition));
						
//						sheet.setAutoFilter(new CellRangeAddress(0,0,0,0));
						sheet.createFreezePane(0, headerPosition);
						line = null;
					}
					
					
					//Set ZoomIn to 85%
					sheet.setZoom(17,20);
					
					// Applying column width
                    for (int cellnum=0; cellnum<cellCount; cellnum++) {
                            int width = (int) ((maxNoChar.get(cellnum) * 7 + 5)/7 * 256)/256;
                            sheet.setColumnWidth(cellnum, width * 256);
                    }
                    reader.close();
				}
				catch (Exception e) {
					System.err.println("\nWarning: Exception encountered. Error Message : " + e.getMessage());
					e.printStackTrace();
					System.exit(1);
				} finally {
					reader = null;
				}
			}
			
			// Write the output to a file
			fileOut = new FileOutputStream(OutputFileName);
			if (FileType.equals("xls")) {
				wb.write(fileOut);
			} else if (FileType.equals("xlsx")) {
				wbx.write(fileOut);
				wbx.dispose();
			}

			fileOut.flush();
			bufferedReader.close(); // closing/flushing bufferedReader
			System.out.println(OutputFileName + " generated and formatted successfully\n");
			
			//For password protection
//			if ( !Zip_password.equalsIgnoreCase("NA") ) {
//				  if(!Zip_password.equals(null)){
//			        // Zip and Password Protect report
//			        System.out.println("Zipping report...");
//			        ZipAndProtectReport zipProtectObject = new ZipAndProtectReport();
//			        String zippedFilename = zipProtectObject.ZipAndProtectMethod(OutputFileName,Zip_password);
//			        System.out.println("Report Zipped and Protected successfully : " + zippedFilename);
//				  }
//			}

		}

		catch (ReadOnlyBufferException e1) {
			System.err.println("\nRead Only Buffer Exception encountered. Error Message : " + e1.getMessage());
			System.exit(1);
		} catch (ArrayIndexOutOfBoundsException e2) {
			System.err.println("\nInput File (Parameter 1) is not in correct format. Expected Format:- <CSV File Name>|<Worksheet Name>. Error Message : " + e2.getMessage());
			System.exit(1);
		} catch (NullPointerException e3) {
			System.err.println("\nNull Pointer Exception encountered. Error Message : " + e3.getMessage());
			System.exit(1);
		} catch (NoClassDefFoundError e4) {
			System.err.println("\nIssue with $CLASSPATH definition. Error Message : " + e4.getMessage());
			System.exit(1);
		} catch (Exception e5) {
			System.err.println("\nException encountered. Error Message : " + e5.getMessage());
			e5.printStackTrace();
			System.exit(1);
		} finally {
			row = null;
			sheet  = null;
			wb = null;
			wbx = null;
			fileOut.close();
		}
	}
	
	public Cell setCellValueAndStyle (Cell cell, CreationHelper helper, int flag, String cellFrmtValue, String cellValue) throws ParseException, NumberFormatException {
		
		boolean strCheck = true;
		SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
		SimpleDateFormat formatter2 = new SimpleDateFormat("yyyy-MM-dd");
		SimpleDateFormat timestampSDF = new SimpleDateFormat("dd/MM/yyyy hh:mm");
		SimpleDateFormat timeFormatter = new SimpleDateFormat("HH:mm:ss");
		Date date = null;

		String str = helper.createRichTextString(cellValue).toString();
		Datatypes currentDay = Datatypes.valueOf(cellFrmtValue); /* enum 'Datatypes' created in the beginning, 
																    all new datatypes need to be added there */
		// Check for nulls and empty strings
		if ( str.equals(null) || str.trim().equals("")) {
			strCheck = true; 
		} else {
			strCheck = false;
		}
		
		try {
		if (flag == 1) {
			switch(currentDay) {
			   case STRING :
				   if ( strCheck == false) {
					   cell.setCellValue(helper.createRichTextString(cellValue));
					   cell.setCellType(Cell.CELL_TYPE_STRING);
					   temp = cell.getRichStringCellValue().length()+6; // +6 for width adjustment to avoid overlap
				   }
				   break;
			   case NUMBER :
				   if ( strCheck == false) {
					   cell.setCellValue(Double.parseDouble(str));
					   cell.setCellType(Cell.CELL_TYPE_NUMERIC);
					   temp = str.length()+6;
				   }
				   break;
			   case DATE :
				   if ( strCheck == false) {
					   date = formatter.parse(str);
					   cell.setCellValue(date);
					   cell.setCellStyle(CreateStyles.dateStyle);
					   temp = str.length()+6;
				   } 
				   break;
			   case DATE2 :
				   if ( strCheck == false) {
					   date = formatter2.parse(str);
					   cell.setCellValue(date);
					   cell.setCellStyle(CreateStyles.date2Style);
					   temp = str.length()+6;
				   } 
				   break;
			   case CURRENCY :
				   if ( strCheck == false) {
					   cell.setCellValue(Double.parseDouble(str.substring(1).trim()));
					   cell.setCellStyle(CreateStyles.currencyStyle);
					   temp = str.length()+6;
				   }
				   break;
			   case TIMESTAMP :
				   if ( strCheck == false) {
					   date = timestampSDF.parse(str);
					   cell.setCellValue(date);
					   cell.setCellStyle(CreateStyles.timestampStyle);
					   temp = str.length()+6;
				   } 
				   break;
			   case TIME :
				   if ( strCheck == false) {
					   date = timeFormatter.parse(str);
					   cell.setCellValue(date);
					   cell.setCellStyle(CreateStyles.timeStyle);
					   temp = str.length()+6;
				   } 
				   break;
			   case NUMBERC1 :
				   if ( strCheck == false) {
					   cell.setCellValue(Double.parseDouble(str));
					   cell.setCellStyle(CreateStyles.Style_number_custom_1);
					   temp = str.length()+6;
				   }
				   break;
			   default : 
				   System.out.println("Invalid Cell Format: " + cellFrmtValue);
			}
		} else {
			cell.setCellType(Cell.CELL_TYPE_STRING);	
			cell.setCellValue(helper.createRichTextString(cellValue));
			temp = cell.getRichStringCellValue().length()+6;
		}
		} 
		catch (Exception e) {
			try {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(helper.createRichTextString(cellValue));
				temp = cell.getRichStringCellValue().length()+6;
			} catch (Exception e1) {
				e1.printStackTrace();
				System.exit(1);
			}
		}
		return cell;
	}
}
