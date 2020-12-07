import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

public class HSSFReadWrite {
	
	public String fileName;
	
	public HSSFReadWrite(String newFileName) {
		this.fileName = newFileName;
	}
	
	public void processSheet() throws Exception {
		try (HSSFWorkbook wb = HSSFReadWrite._readFile(this.fileName)) { 
			System.out.println("hssf");
			_readExcel(wb);
		}
	}
	
	private void _readExcel(HSSFWorkbook wb) throws IOException {
		
		/*
		 	==========================================
		  	Only working with the first sheet
			==========================================
		 */
		
		Sheet sheet = wb.getSheetAt(0);
		DataFormatter dataFormatter = new DataFormatter();
		
		String[] keys = new String[sheet.getLastRowNum()+1];
		String[][] dictionaries = new String[sheet.getRow(0).getLastCellNum()+1][sheet.getLastRowNum()+1];
		
		int rowNb = 0;
		int colKey = 0;
		
		/*
		 * FIX THIS PART
		 */
		
		for (Row row: sheet) {
			int colNb = 0;
			
			for (Cell cell: row) {
				if (dataFormatter.formatCellValue(cell).toLowerCase().equals("key")) {
					colKey = colNb;
					System.out.println(colKey);
				}
				
				if (dataFormatter.formatCellValue(cell).length() == 2) {
					System.out.println("lang");
				}
				
				colNb++;
			}
			rowNb++;
		}
		
		/*
		for (int i = 0; i < dictionaries.length; i++) {
			_writeToPropertiesFile(keys,dictionaries[i]);
		}
		*/
		wb.close();
	}
	
	private void _writeToPropertiesFile(String[] keys, String[] values) throws IOException {
		Properties props = new Properties();
		
		for (int i = 0; i < keys.length; i++) {
			if (keys[i] != null && keys[i] != "") {
				props.setProperty(keys[i], "fix");
			}
		}
		
		FileOutputStream fileopts = new FileOutputStream(new File(getFileNameBase(this.fileName) + "-" + "fr" + ".properties"));
		
		props.store(fileopts, null);
		
		fileopts.close();
	}
	
	private static HSSFWorkbook _readFile(String fileName) throws IOException { 
		try (POIFSFileSystem fs = new POIFSFileSystem(new File(fileName))) {
			return new HSSFWorkbook(fs);        // NOSONAR - should not be closed here
		}
	}
	
	private static String getFileNameBase(String fileName) {
		String[] arrOfStr = fileName.split("\\.(?=[^\\.]+$)");
		
		return arrOfStr[0];
	}
}
