import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.commons.compress.compressors.FileNameUtil;

public class Program {

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
				HSSFReadWrite hssfRW = new HSSFReadWrite(fileName);
				hssfRW.processSheet();
				break;
			case "xlsx":
				XSSFReadWrite xssfRW = new XSSFReadWrite(fileName);
				xssfRW.processSheet(xssfRW.fileName);
				break;
			default:
				System.err.println(fileName + " : Not valid extension must be .xls or .xlsx");
				break;
			}
		}
	}
	
	// Methods that returns the extension
	private static String getFileExtension(String fileName) {
		String extension = "";
		
		int i = fileName.lastIndexOf('.');
		if (i > 0) {
		    extension = fileName.substring(i+1);
		}
		
		return extension;
	}
}
