package ExcelConversion;

import ExcelConversion.ConvertToExcel;
import ExcelConversion.ConvertAndFormatExcel;

import java.io.FileOutputStream;
import java.io.IOException;

/* * * * * * * * * * * * * * * * * * * * * * * * * */
/*  	Developer : Richa Verma					   */
/*  	Date : 06/01/2017 						   */
/* * * * * * * * * * * * * * * * * * * * * * * * * */

public class CSVToEXCEL {
	
	public static String reportName = "";
	static FileOutputStream fileOut = null;

	public static void main(String args[]) throws IOException {

		System.setProperty("java.io.tmpdir", System.getProperty("ROOT") + "/temp");
		
		ConvertToExcel cte = new ConvertToExcel();
		ConvertAndFormatExcel cfe = new ConvertAndFormatExcel();
		
		try {
			if(args.length == 3) {
				cte.doConversion(args[0], args[1], args[2], "NA"); 
			}else if(args.length == 4) {
				cte.doConversion(args[0], args[1], args[2], args[3]);
			}else if(args.length == 5) { 
				cfe.convertAndFormat(args[0], args[1], args[2], args[3], args[4]);
			}else if(args.length == 6) {
				reportName = args[5];
				cfe.convertAndFormat(args[0], args[1], args[2], args[3], args[4]);
			}else {
				System.out.println("Wrong number of arguments passed");
				System.exit(1);
			}
		}

		catch (ArrayIndexOutOfBoundsException e) {
			System.out.println("Input parameters are not passed correctly in JAVA call. Parameters should be:- " +
					"\n 1. File List Name (CSV Filename | Worksheet Name <<| Report Format File Name//reqd with formatting only//>>) \n "
					+ "2. Output File Name \n 3. Output File Format \n 4. Banner Length \n 5. Header Position or (Header Position and password delimited by #) "
					+ "\n 6. Report Name (Applicable for STD-AF only)");
			System.exit(1);
		} catch (Exception e) {
			System.err.println("Other issues - " + e.getMessage());
			System.exit(1);
		} 
	}
}
