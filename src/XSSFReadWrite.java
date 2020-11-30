import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XSSFReadWrite {
	
	public String fileName;
	
	public XSSFReadWrite(String newFileName) {
		this.fileName = newFileName;
	}
	
	public void processSheet(String fileName) throws Exception {
		try (XSSFWorkbook wb = XSSFReadWrite.readFile(fileName)) {
			System.out.println("xssf");
		}
	}
	
	private static XSSFWorkbook readFile(String fileName) throws Exception { 
		try (OPCPackage pkg = OPCPackage.open(fileName, PackageAccess.READ)) {
			return new XSSFWorkbook(pkg);      // NOSONAR - should not be closed here
		}
	}
}
