import java.io.File;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class HSSFReadWrite extends ReadWrite<HSSFWorkbook> {

	public HSSFReadWrite(String newFileName) {
		super(newFileName);
	}

	@Override
	public void processSheet() throws Exception {
		try (HSSFWorkbook wb = readFile(this.fileName)) { 
			_readExcel(wb);
		}
	}

	@Override
	public HSSFWorkbook readFile(String fileName) throws IOException {
		try (POIFSFileSystem fs = new POIFSFileSystem(new File(fileName))) {
			return new HSSFWorkbook(fs);        // NOSONAR - should not be closed here
		}
	}

}
