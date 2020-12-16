import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.*;

public class HSSFReadWrite {
	
	public String fileName;
	
	public HSSFReadWrite(String newFileName) {
		this.fileName = newFileName;
	}
	
	public void processSheet() throws Exception {
		try (HSSFWorkbook wb = HSSFReadWrite._readFile(this.fileName)) { 
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
		
		int colKey = 0; // Column number where are the keys
		boolean lookingForHeader = true; // Flag to know if we're still looking for the header row
		
		Map<Integer, Language> translationsMap = new HashMap<>();
		
		//===============================================
		// We loop through the sheet
		//===============================================
		Iterator<Row> rowIterator = sheet.iterator();
		
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			
			Iterator<Cell> cellIterator = row.cellIterator();
			
			if (lookingForHeader && !checkIfRowIsEmpty(row)) { // Read Header
				while(cellIterator.hasNext()) {
					
					Cell cell = cellIterator.next();
					
					if (dataFormatter.formatCellValue(cell).toLowerCase().equals("key")) { // Looking for the header column
						colKey = cell.getColumnIndex();
					}
					
					if (dataFormatter.formatCellValue(cell).length() == 2) { // Looking for all languages
						translationsMap.put(cell.getColumnIndex(), new Language(dataFormatter.formatCellValue(cell)));
					}
				}
				
				lookingForHeader = false;
			} else { // Read Translations
				while(cellIterator.hasNext()) {
					
					Cell cell = cellIterator.next();
					
					if (translationsMap.containsKey(cell.getColumnIndex())) {
						translationsMap.get(cell.getColumnIndex()).addTranslation(dataFormatter.formatCellValue(row.getCell(colKey)), dataFormatter.formatCellValue(cell));
					}
				}
			}
		}
		//===============================================
		
		_writeToPropertiesFile(translationsMap); // We call the method to write the properties file(s).
		wb.close();
	}
	
	private void _writeToPropertiesFile(Map<Integer, Language> translationsMap) throws IOException {
		for (Map.Entry<Integer, Language> entry : translationsMap.entrySet()) {
			Properties props = new Properties();
			
			for (Map.Entry<String, String> string : entry.getValue().translations.entrySet()) {
				props.setProperty(string.getKey(), string.getValue());
			}
			
			FileOutputStream fileopts = new FileOutputStream(new File(getFileNameBase(this.fileName) + "-" + entry.getValue().lang + ".properties"));
			
			props.store(fileopts, null);
			
			fileopts.close();
			
		}
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
	
	private boolean checkIfRowIsEmpty(Row row) {
	    if (row == null) {
	        return true;
	    }
	    if (row.getLastCellNum() <= 0) {
	        return true;
	    }
	    for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
	        Cell cell = row.getCell(cellNum);
	        if (cell != null && cell.getCellType() != CellType.BLANK) {
	            return false;
	        }
	    }
	    return true;
	}
	
	
}
