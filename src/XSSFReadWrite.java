import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XSSFReadWrite extends ReadWrite<XSSFWorkbook> {

	public XSSFReadWrite(String newFileName) {
		super(newFileName);
	}
	
	@Override
	public void processSheet() throws Exception {
		try (XSSFWorkbook wb = readFile(this.fileName)) { 
			_readExcel(wb);
		}
	}

	@Override
	public XSSFWorkbook readFile(String fileName) throws Exception {
		try (OPCPackage pkg = OPCPackage.open(fileName, PackageAccess.READ)) {
			return new XSSFWorkbook(pkg);
		}
	}

}
