package reader;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import VKModel.VKModel;

public class XSLXReader {
	
	XSSFWorkbook doc;
	XSSFSheet sheet;
	
	public XSLXReader(String path, int sheetNr) throws InvalidFormatException, IOException{
		File file = new File(path);

		doc = new XSSFWorkbook(file);
		sheet = doc.getSheetAt(sheetNr);
	}
	
	public VKModel readRow(int rowNr, int VName, int NName, int Anrede, int Firma, int Abteilung, int Position, int TelBuero, int TelMobil, int TelFax, int AdStraﬂe, int AdPLZ, int AdStadt, int AdLand, int Sprache, int Notiz, int Veranstaltung, int Email, int Website){
		Row row = sheet.getRow(rowNr);
		VKModel vk = new VKModel();
		
		vk.setVName(row.getCell(VName).getStringCellValue());
		vk.setNName(row.getCell(NName).getStringCellValue());
		vk.setAnrede(row.getCell(Anrede).getStringCellValue());
		vk.setFirma(row.getCell(Firma).getStringCellValue());
		vk.setAbteilung(row.getCell(Abteilung).getStringCellValue());
		vk.setPosition(row.getCell(Position).getStringCellValue());
		vk.setTelBuero(row.getCell(TelBuero).getStringCellValue());
		vk.setTelMobil(row.getCell(TelMobil).getStringCellValue());
		vk.setTelFax(row.getCell(TelFax).getStringCellValue());
		vk.setAdStraﬂe(row.getCell(AdStraﬂe).getStringCellValue());
		vk.setAdPLZ(row.getCell(AdPLZ).getStringCellValue());
		vk.setAdStadt(row.getCell(AdStadt).getStringCellValue());
		vk.setAdLand(row.getCell(AdLand).getStringCellValue());
		vk.setSprache(row.getCell(Sprache).getStringCellValue());
		vk.setNotiz(row.getCell(Notiz).getStringCellValue());
		vk.setVeranstaltung(row.getCell(Veranstaltung).getStringCellValue());
		vk.setEmail(row.getCell(Email).getStringCellValue());
		vk.setWebsite(row.getCell(Website).getStringCellValue());
		
		return vk;
	}

}
