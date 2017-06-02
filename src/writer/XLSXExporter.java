package writer;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import VKModel.VKModel;

public class XLSXExporter {
	
	XSSFWorkbook doc;
	XSSFSheet sheet;
	
	public XLSXExporter(){
		doc = new XSSFWorkbook();
		sheet = doc.createSheet("Import");
	}
	
	public void addVK(VKModel model){
		
	}
	
	public void exportAsFile(String path) throws IOException{
		FileOutputStream fileOut = new FileOutputStream(path);
	    doc.write(fileOut);
	    fileOut.close();
	}
	

}
