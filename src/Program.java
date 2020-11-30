import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.commons.compress.compressors.FileNameUtil;

public class Program {

	public static void main(String[] args) throws Exception {
		
		if (args.length < 1) {
			System.err.println("At least one argument expected");
			return;
		}
		
		for (String fileName : args) {
			String fileExtension = getFileExtension(fileName);
			
			switch(fileExtension) {
			
			case "xls":
				HSSFReadWrite hssfRW = new HSSFReadWrite(fileName);
				hssfRW.processSheet(hssfRW.fileName);
				break;
			case "xlsx":
				XSSFReadWrite xssfRW = new XSSFReadWrite(fileName);
				xssfRW.processSheet(xssfRW.fileName);
				break;
			default:
				System.err.println("Not valid extension must be .xls or .xlsx");
				break;
			}
		}
	}
	
	private static String getFileExtension(String fileName) {
		String extension = "";
		
		int i = fileName.lastIndexOf('.');
		if (i > 0) {
		    extension = fileName.substring(i+1);
		}
		
		return extension;
	}
}
