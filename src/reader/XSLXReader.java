package reader;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import VKModel.VKModel;

public class XSLXReader {
	
	XSSFWorkbook doc;
	XSSFSheet sheet;
	public int rowCount;
	
	public XSLXReader(String path, int sheetNr) throws InvalidFormatException, IOException{
		File file = new File(path);
		doc = new XSSFWorkbook(file);
		sheet = doc.getSheetAt(sheetNr);
		rowCount = sheet.getLastRowNum();
		doc.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);
	}
	
	public VKModel readRow(int rowNr, Integer VName, Integer NName, Integer Anrede, Integer Firma, Integer Abteilung, Integer Position, Integer TelBuero, Integer TelMobil, Integer TelFax, Integer AdStraﬂe, Integer AdPLZ, Integer AdStadt, Integer AdLand, Integer Sprache, Integer Notiz, Integer Veranstaltung, Integer Email, Integer Website){
		Row row = sheet.getRow(rowNr);
		VKModel vk = new VKModel();
		if(VName!=null){
			vk.setVName(row.getCell(VName).getStringCellValue());
		} else {vk.setVName("");}
		
		if(NName!=null){
			vk.setNName(row.getCell(NName).getStringCellValue());
		} else {vk.setNName("");}
		
		if(Anrede!=null){
			vk.setAnrede(row.getCell(Anrede).getStringCellValue());
		} else {vk.setAnrede("");}
		
		if(Firma!=null){
			vk.setFirma(row.getCell(Firma).getStringCellValue());
		} else {vk.setFirma("");}
		
		if(Abteilung!=null){
			vk.setAbteilung(row.getCell(Abteilung).getStringCellValue());
		} else {vk.setAbteilung("");}
		
		if(Position!=null){
			vk.setPosition(row.getCell(Position).getStringCellValue());
		} else {vk.setPosition("");}
		
		if(TelBuero!=null){
			row.getCell(TelBuero).setCellType(Cell.CELL_TYPE_STRING);
			vk.setTelBuero(row.getCell(TelBuero).getStringCellValue());
		} else {vk.setTelBuero("");}
		
		if(TelMobil!=null){
			row.getCell(TelMobil).setCellType(Cell.CELL_TYPE_STRING);
			vk.setTelMobil(row.getCell(TelMobil).getStringCellValue());
		} else {vk.setTelMobil("");}
		
		if(TelFax!=null){
			row.getCell(TelFax).setCellType(Cell.CELL_TYPE_STRING);
			vk.setTelFax(row.getCell(TelFax).getStringCellValue());
		} else {vk.setTelFax("");}
		
		if(AdStraﬂe!=null){
			vk.setAdStraﬂe(row.getCell(AdStraﬂe).getStringCellValue());
		} else {vk.setAdStraﬂe("");}
		
		if(AdPLZ!=null){
			row.getCell(AdPLZ).setCellType(Cell.CELL_TYPE_STRING);
			vk.setAdPLZ(row.getCell(AdPLZ).getStringCellValue());
		} else {vk.setAdPLZ("");}
		
		if(AdStadt!=null){
			vk.setAdStadt(row.getCell(AdStadt).getStringCellValue());
		} else {vk.setAdStadt("");}
		
		if(AdLand!=null){
			vk.setAdLand(row.getCell(AdLand).getStringCellValue());
		} else {vk.setAdLand("");}
		
		if(Sprache!=null){
			vk.setSprache(row.getCell(Sprache).getStringCellValue());
		} else {vk.setSprache("");}
		
		if(Notiz!=null){
			vk.setNotiz(row.getCell(Notiz).getStringCellValue());
		} else {vk.setNotiz("");}
		
		if(Veranstaltung!=null){
			vk.setVeranstaltung(row.getCell(Veranstaltung).getStringCellValue());
		} else {vk.setVeranstaltung("");}
		
		if(Email!=null){
			vk.setEmail(row.getCell(Email).getStringCellValue());
		} else {vk.setEmail("");}
		
		if(Website!=null){
			vk.setWebsite(row.getCell(Website).getStringCellValue());
		} else {vk.setWebsite("");}
		
		return vk;
	}

}
