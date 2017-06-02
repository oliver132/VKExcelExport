package writer;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import VKModel.VKModel;

public class XLSXExporter {
	
	XSSFWorkbook doc;
	XSSFSheet sheet;
	int currentLine;
	
	public XLSXExporter(){
		doc = new XSSFWorkbook();
		sheet = doc.createSheet("Import");
		currentLine=1;
	}
	
	public void addVK(VKModel model){
		Row row = sheet.createRow(currentLine);

		row.createCell(1).setCellValue(model.getVName());
		row.createCell(2).setCellValue(model.getNName());
		row.createCell(3).setCellValue(model.getFirma());
		row.createCell(4).setCellValue(model.getAbteilung());
		row.createCell(5).setCellValue(model.getPosition());
		row.createCell(6).setCellValue(model.getAnrede());
		row.createCell(7).setCellValue(model.getAdStraﬂe());
		row.createCell(8).setCellValue(model.getAdPLZ());
		row.createCell(9).setCellValue(model.getAdStadt());
		row.createCell(10).setCellValue(model.getAdLand());
		row.createCell(11).setCellValue(model.getSprache());
		row.createCell(12).setCellValue(model.getTelBuero());
		row.createCell(13).setCellValue(model.getTelFax());
		row.createCell(14).setCellValue(model.getTelMobil());
		row.createCell(15).setCellValue(model.getNotiz());
		row.createCell(16).setCellValue(model.getVeranstaltung());
		row.createCell(17).setCellValue(model.getEmail());
		row.createCell(18).setCellValue(model.getWebsite());
		
		currentLine+=1;
		
	}
	
	public void exportAsFile(String path) throws IOException{
		FileOutputStream fileOut = new FileOutputStream(path);
	    doc.write(fileOut);
	    fileOut.close();
	}
	

}
