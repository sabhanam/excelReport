package com.ramSabhanam.excelReport;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.ramSabhanam.Database.IDatabase;
import com.ramSabhanam.Database.MicrosoftExcel;

public class ExcelReport {

	IDatabase database = null;
	
	public static void main(String... args) {
		
		Map<String,String> data = new HashMap<String,String>();
		data.put("Column1", "asdas");
		data.put("Sairam", "asd");
		data.put("Thirumala", "asx");
		data.put("ExecutionStatus", "PASSED");
		data.put("qwsd", "sdwd");
		
		
		ExcelReport excelReport = new ExcelReport("xlsx","/Users/sunnysabhanam/Desktop/Report");
		excelReport.flushExcel(data);
		excelReport.dispose();
		
	}
	
	public ExcelReport(String databaseType, String databaseLocation) {
		
		if(databaseType.equals("xlsx") || databaseType.equals("xls")) {
			
			database = new MicrosoftExcel(databaseType, databaseLocation);
			
		}
		
	}

	public void flushExcel(Map<String, String> data) {
		
		database.flushExcel(data);
		
	}

	public void flushExcel(List<Map<String, String>> data) {
		
		database.flushExcel(data);
		
	}

	public void dispose() {
		
		if(database != null)
			database.dispose();
		
	}
	
}
