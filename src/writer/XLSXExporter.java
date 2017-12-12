package writer;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
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
	
	public XLSXExporter() throws IOException{
		FileInputStream fIPS= new FileInputStream("importVorlage.xlsx");
		doc =  new XSSFWorkbook(fIPS);
		sheet = doc.getSheetAt(0);
		currentLine=1;
	}
	
	public void addVK(VKModel model){
		Row row = sheet.createRow(currentLine);

		row.createCell(0).setCellValue("dat");
		row.createCell(2).setCellValue("1");
		row.createCell(66).setCellValue("1");
		row.createCell(8).setCellValue(model.getVName());
		row.createCell(9).setCellValue(model.getNName());
		row.createCell(67).setCellValue(model.getFirma());
		row.createCell(13).setCellValue(model.getAbteilung());
		row.createCell(14).setCellValue(model.getPosition());
		row.createCell(17).setCellValue(model.getAnrede());
		row.createCell(70).setCellValue(model.getAdStraﬂe());
		row.createCell(68).setCellValue(model.getAdPLZ());
		row.createCell(69).setCellValue(model.getAdStadt());
		
		//land
		if (model.getAdLand().equals("Germany")){
			row.createCell(73).setCellValue("Deutschland");
			
		} else {
			row.createCell(73).setCellValue(model.getAdLand());
		}
		
		
		
		row.createCell(25).setCellValue(model.getTelBuero());
		row.createCell(31).setCellValue(model.getTelFax());
		row.createCell(27).setCellValue(model.getTelMobil());
		row.createCell(39).setCellValue(model.getNotiz());
		row.createCell(118).setCellValue(model.getVeranstaltung());
		row.createCell(120).setCellValue(model.getVeranstaltung());
		row.createCell(29).setCellValue(model.getEmail());
		row.createCell(87).setCellValue(model.getWebsite());
		
		if (model.getSprache().equals("")) {
			if(numberIsInDACH(model.getTelBuero())||numberIsInDACH(model.getTelMobil())||numberIsInDACH(model.getTelFax())){
				row.createCell(38).setCellValue("deutsch");
			} else {
				row.createCell(38).setCellValue("english");
				
			}
		} else {
			row.createCell(38).setCellValue(model.getSprache());
		}
		//Sprache je nach TelVorwahl

		
		currentLine+=1;
		
	}
	
	public void exportAsFile(String name) throws IOException{
		FileOutputStream fileOut = new FileOutputStream(System.getProperty("user.home") + "/Desktop/"+name);
	    doc.write(fileOut);
	    fileOut.close();
	}
	
	public boolean numberIsInDACH(String telnr){
		if(telnr.contains("+49")||telnr.contains("+41")||telnr.contains("+43")||telnr.contains("0041")||telnr.contains("0043")||telnr.contains("0049")){
			return true;
		}
		return false;
	}
}
