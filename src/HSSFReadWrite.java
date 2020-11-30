import java.io.File;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class HSSFReadWrite {
	
	public String fileName;
	
	public HSSFReadWrite(String newFileName) {
		this.fileName = newFileName;
	}
	
	public void processSheet(String fileName) throws Exception {
		try (HSSFWorkbook wb = HSSFReadWrite.readFile(fileName)) { 
			System.out.println("hssf");
		}
	}
	
	private static HSSFWorkbook readFile(String fileName) throws IOException { 
		try (POIFSFileSystem fs = new POIFSFileSystem(new File(fileName))) {
			return new HSSFWorkbook(fs);        // NOSONAR - should not be closed here
		}
	}
}
