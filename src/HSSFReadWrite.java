import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class HSSFReadWrite {
	
	public String fileName;
	
	public HSSFReadWrite(String newFileName) {
		this.fileName = newFileName;
	}
	
	public void processSheet(String fileName) throws Exception {
		try (HSSFWorkbook wb = HSSFReadWrite._readFile(fileName)) { 
			System.out.println("hssf");
			
			_writeToPropertiesFile();
		}
	}
	
	private void _writeToPropertiesFile() throws IOException {
		Properties props = new Properties();
		
		props.setProperty("key", "value");
		
		FileOutputStream fileopts = new FileOutputStream(new File("propsxls.properties"));
		
		props.store(fileopts, null);
		
		fileopts.close();
	}
	
	private static HSSFWorkbook _readFile(String fileName) throws IOException { 
		try (POIFSFileSystem fs = new POIFSFileSystem(new File(fileName))) {
			return new HSSFWorkbook(fs);        // NOSONAR - should not be closed here
		}
	}
}
