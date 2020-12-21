import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.commons.compress.compressors.FileNameUtil;

public class OldProgram {

	public static void main(String[] args) throws Exception {
		
		/*
		 * You must pass arguments as the filename you wanna read.
		 * It will automatically generate the .properties file.
		 */
		
		// Checks if the user did pass an argument
		if (args.length < 1) {
			System.err.println("At least one argument expected");
			return;
		}
		
		// for each arguments passed check the extension to see if it's valid or not (xls or xlsx) if no return an error.
		for (String fileName : args) {
			String fileExtension = getFileExtension(fileName);
			
			switch(fileExtension) {
			
			case "xls":
				System.out.println("xls file detected");
				OldHSSFReadWrite hssfRW = new OldHSSFReadWrite(fileName);
				hssfRW.processSheet();
				break;
			case "xlsx":
				System.out.println("xlsx file detected");
				OldXSSFReadWrite xssfRW = new OldXSSFReadWrite(fileName);
				xssfRW.processSheet(xssfRW.fileName);
				break;
			default:
				System.err.println(fileName + " : Not valid extension must be .xls or .xlsx");
				break;
			}
		}
	}
	
	/**
	 * Returns the extension of the file.
	 * @param fileName	a complete file name (example: test.txt). 
	 * @return			the extension of the file (example: txt).
	 */
	private static String getFileExtension(String fileName) {
		String extension = "";
		
		int i = fileName.lastIndexOf('.');
		if (i > 0) {
		    extension = fileName.substring(i+1);
		}
		
		return extension;
	}
}
