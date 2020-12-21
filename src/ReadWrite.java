import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Class that permit to read excel files (xls and xlsx), and create properties files for each language.
 * @author brian.grn@eduge.ch
 *
 * @param <T>	Generic type depending on the file extension HSSFWorkbook for xls and XSSFWorkbook for xlsx.
 */
public abstract class ReadWrite<T extends Workbook> {
public String fileName;
	
	public ReadWrite(String newFileName) {
		this.fileName = newFileName;
	}
	
	public abstract void processSheet() throws Exception;
	
	public abstract T readFile(String fileName) throws IOException, Exception;
	
	/**
	 * Read an excel, loops inside it to get all data and then calls the method {@link #_writeToPropertiesFile(Map)} 
	 * to write the properties files.
	 * @param wb			the workbook of our file containing all excel sheets.
	 * @throws IOException
	 */
	@SuppressWarnings("unused")
	protected void _readExcel(T wb) throws IOException {
		
		/*
		 *	==========================================
		 * 	Only working with the first sheet
		 *	==========================================
		 */
		
		Sheet sheet = wb.getSheetAt(0);
		DataFormatter dataFormatter = new DataFormatter();
		
		int colKey = 0; // Column number where are the keys
		boolean lookingForHeader = true; // Flag to know if we're still looking for the header row
		
		Map<Integer, Language> translationsMap = new HashMap<>();
		
		//===============================================
		// We loop through the sheet
		//===============================================
		
		System.out.println("Looping inside the sheet");
		
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
	
	/**
	 * Creates properties file depending on the number of language in the @param translationsMap. It loop inside the map in the language
	 * to get the keys and values.
	 * @param translationsMap	map of column number to language.
	 * @throws IOException
	 * @see Language
	 */
	private void _writeToPropertiesFile(Map<Integer, Language> translationsMap) throws IOException {
		System.out.println("Writing properties file(s).");
		for (Map.Entry<Integer, Language> entry : translationsMap.entrySet()) {
			Properties props = new Properties();
			
			for (Map.Entry<String, String> string : entry.getValue().translations.entrySet()) {
				props.setProperty(string.getKey(), string.getValue());
			}
			
			FileOutputStream fileopts = new FileOutputStream(new File(getFileNameBase(this.fileName) + "-" + entry.getValue().lang + ".properties"));
			System.out.println(getFileNameBase(this.fileName) + "-" + entry.getValue().lang + ".properties created");
			
			props.store(fileopts, null);
			
			fileopts.close();
		}
		
		System.out.println(this.fileName + " finished");
	}
	
	/**
	 * Returns the base of the file name it removes the extension.
	 * @param fileName	a complete file name (example: test.txt).
	 * @return			the base of the file name (example test).
	 */
	private String getFileNameBase(String fileName) {
		String[] arrOfStr = fileName.split("\\.(?=[^\\.]+$)");
		
		return arrOfStr[0];
	}
	
	/**
	 * Checks if the current is empty or not.
	 * @param row	row of the sheet.
	 * @return		True if the row is empty else False.
	 */
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
