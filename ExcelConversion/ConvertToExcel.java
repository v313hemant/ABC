package ExcelConversion;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.nio.ReadOnlyBufferException;
import java.util.ArrayList;
import java.util.List;
import au.com.bytecode.opencsv.CSVReader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/* * * * * * * * * * * * * * * * * * * * * * * * * */
/*  	Developer : Richa Verma					   */
/*  	Date : 28/12/2016 						   */
/* * * * * * * * * * * * * * * * * * * * * * * * * */

public class ConvertToExcel {

	static HSSFWorkbook wb = new HSSFWorkbook();
	static SXSSFWorkbook wbx = new SXSSFWorkbook();
	static FileOutputStream fileOut = null;
	
	public void doConversion(String Flist, String OutFname, String Ftype , String Zip_password) throws IOException {

		// Declaring variables
		List<String> lines = new ArrayList<String>();
				
		Cell cell = null;
		Row row = null;
		Sheet sheet = null;
		CreationHelper helper = null;
		
		int temp = 15;
		ArrayList<Integer> maxNoChar = new ArrayList<Integer>();
		
		String InputString = null;
		
		String FileName = null;
		String infile = null;
		String WorksheetName = null;
		String OutputFileName = "";
		String[] parts = null;
		String[] line = null;
		
		try {
			FileReader fileReader = new FileReader(Flist);
			
			BufferedReader bufferedReader = new BufferedReader(fileReader);
			
			String FOutName = OutFname;
			String FileType = Ftype.toLowerCase();
			
			System.out.println(FOutName+"."+FileType);
			
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
				System.out.println("Input CSV Name : " + FileName + "\nWorksheet Name : " + WorksheetName);
				
				if (FileType.equals("xls")) {
					helper = wb.getCreationHelper();
					sheet = wb.createSheet(WorksheetName); // Create XLS Worksheet
				} else if (FileType.equals("xlsx")) {
					helper = wbx.getCreationHelper();
					sheet = wbx.createSheet(WorksheetName); // Create XLSX Worksheet
				} else {
					System.out.print("Invalid File Type");
				}
				
				CSVReader reader = null;
				
				try {
					reader = new CSVReader(new FileReader(FileName)); // Reading Input CSV files
					long r = 0;
	
					while ((line = reader.readNext()) != null) {
						row = sheet.createRow((int) r++);
						
						for (int j = 0; j < line.length; j++) {
							cell = row.createCell(j);
							cell.setCellValue(helper.createRichTextString(line[j]));
							temp = cell.getRichStringCellValue().length()+8; // +8 for width adjustment to avoid overlap
							
							// Reading maximum length for each row for column width setting
							if (maxNoChar.size() < line.length) {
								maxNoChar.add(j, temp);
							} else if (temp > maxNoChar.get(j)) {
								maxNoChar.set(j, temp);
							}
						}
					}
					line = null;
					
					// Applying column width
					for (int cellnum=0; cellnum<maxNoChar.size(); cellnum++) {
						int width = (int) ((maxNoChar.get(cellnum) * 7 + 5)/7 * 256)/256;
                        sheet.setColumnWidth(cellnum, width * 256);
					}
					reader.close();
                } catch (Exception e) {
					System.err.println("\nWarning: Exception encountered. Error Message : " + e.getMessage());
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
			System.out.println(OutputFileName + " generated successfully\n");
			
			  if ( !Zip_password.equalsIgnoreCase("NA") ) {
				  if(!Zip_password.equals(null)){
			        // Zip and Password Protect report
			        System.out.println("Zipping report...");
			        ZipAndProtectReport zipProtectObject = new ZipAndProtectReport();
			        String zippedFilename = zipProtectObject.ZipAndProtectMethod(OutputFileName,Zip_password);
			        System.out.println("Report Zipped and Protected successfully : " + zippedFilename);
				  }
			}
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
			System.exit(1);
		}  finally {
			row = null;
			sheet  = null;
			wb = null;
			wbx = null;
			fileOut.close();
		}
	}
}
