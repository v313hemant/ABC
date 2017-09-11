package ExcelConversion;

import java.io.File;
import java.util.ArrayList;

import net.lingala.zip4j.core.ZipFile;
import net.lingala.zip4j.exception.ZipException;
import net.lingala.zip4j.model.ZipParameters;
import net.lingala.zip4j.util.Zip4jConstants;


public class ZipAndProtectReport {

	public String ZipAndProtectMethod (String XLSXfile, String Password){	

		String ZippedFileName = XLSXfile.substring(0, XLSXfile.lastIndexOf('.')) + ".zip";
		
		ZipFile zipFile;
		try {
			zipFile = new ZipFile(ZippedFileName);
		
		ArrayList<File> filesToAdd = new ArrayList<File>();

		filesToAdd.add(new File(XLSXfile));
		
		// Setting parameters for File Zip
        ZipParameters parameters = new ZipParameters();
        parameters.setCompressionMethod(Zip4jConstants.COMP_DEFLATE); 
        parameters.setCompressionLevel(Zip4jConstants.DEFLATE_LEVEL_NORMAL); 
        
        // Password applicable when argument <> NA
        if (!Password.equals("NA")) {
	        parameters.setEncryptFiles(true);
	        parameters.setEncryptionMethod(Zip4jConstants.ENC_METHOD_STANDARD);
        	parameters.setPassword(Password);
        }
        
        zipFile.addFiles(filesToAdd, parameters);
        
		} catch (ZipException e) {
			System.err.println("Unknow exception, Please refer stack trace.");
			System.exit(1);
			e.printStackTrace();
		}
		
		return ZippedFileName;
	}

}

