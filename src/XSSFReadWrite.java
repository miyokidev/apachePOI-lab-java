import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XSSFReadWrite {
	
	public String fileName;
	
	public XSSFReadWrite(String newFileName) {
		this.fileName = newFileName;
	}
	
	public void processSheet(String fileName) throws Exception {
		try (XSSFWorkbook wb = XSSFReadWrite._readFile(fileName)) {
			System.out.println("xssf");
			_writeToPropertiesFile();
		}
	}
	
	private void _writeToPropertiesFile() throws IOException {
		Properties props = new Properties();
		
		props.setProperty("key", "value");
		
		FileOutputStream fileopts = new FileOutputStream(new File("propsxlsx.properties"));
		
		props.store(fileopts, null);
		
		fileopts.close();
	}
	
	private static XSSFWorkbook _readFile(String fileName) throws Exception { 
		try (OPCPackage pkg = OPCPackage.open(fileName, PackageAccess.READ)) {
			return new XSSFWorkbook(pkg);      // NOSONAR - should not be closed here
		}
	}
}
