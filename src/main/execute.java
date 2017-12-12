package main;

import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import VKModel.VKModel;
import reader.XSLXReader;
import writer.XLSXExporter;

public class execute {

	public static void main(String[] args) throws Exception, IOException {
		// TODO Auto-generated method stub
		
		XSLXReader reader = new XSLXReader("C:\\Users\\Oli\\Desktop\\test.xlsx", 0);
		ArrayList<VKModel> vk = new ArrayList<VKModel>();
		
		/*
		 * int rowNr, Integer VName, Integer NName, 
		 * Integer Anrede, Integer Firma, Integer Abteilung, Integer Position, 
		 * Integer TelBuero, Integer TelMobil, Integer TelFax, 
		 * Integer AdStraﬂe, Integer AdPLZ, Integer AdStadt, Integer AdLand, 
		 * Integer Sprache, Integer Notiz, Integer Veranstaltung, Integer Email, Integer Website
		 */
		for (int i=2; i<reader.rowCount+1; i++){
			vk.add(reader.readRow(i, 3, 4, null, 0, 6, 5, 8, 9, null, 10, 11, 12, null, null, 14, null, 7, 25));
		}
		
		XLSXExporter exporter = new XLSXExporter();
		for(VKModel vk_temp : vk){
			exporter.addVK(vk_temp);
		}
		
		exporter.exportAsFile("tabelle1.xlsx");

	}

}
