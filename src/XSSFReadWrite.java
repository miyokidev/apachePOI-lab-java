import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XSSFReadWrite {
	
	public String fileName;
	
	public XSSFReadWrite(String newFileName) {
		this.fileName = newFileName;
	}
	
	public void processSheet(String fileName) throws Exception {
		try (XSSFWorkbook wb = XSSFReadWrite._readFile(fileName)) {
			_readExcel(wb);
		}
	}
	
	private void _readExcel(XSSFWorkbook wb) throws IOException {
		
		/*
		 	==========================================
		  	Only working with the first sheet
			==========================================
		 */
		
		Sheet sheet = wb.getSheetAt(0);
		DataFormatter dataFormatter = new DataFormatter();
		
		int rowNb = 0;
		int colKey = 0;
		boolean lookingForHeader = true;
		
		Map<Integer, Language> translationsMap = new HashMap<>();
		
		
		Iterator<Row> rowIterator = sheet.iterator();
		
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			
			Iterator<Cell> cellIterator = row.cellIterator();
			
			if (lookingForHeader && !checkIfRowIsEmpty(row)) { // Read Header
				while(cellIterator.hasNext()) {
					
					Cell cell = cellIterator.next();
					
					if (dataFormatter.formatCellValue(cell).toLowerCase().equals("key")) {
						colKey = cell.getColumnIndex();
					}
					
					if (dataFormatter.formatCellValue(cell).length() == 2) {
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
		
		_writeToPropertiesFile(translationsMap);
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
	
	private static XSSFWorkbook _readFile(String fileName) throws Exception { 
		try (OPCPackage pkg = OPCPackage.open(fileName, PackageAccess.READ)) {
			return new XSSFWorkbook(pkg);      // NOSONAR - should not be closed here
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
