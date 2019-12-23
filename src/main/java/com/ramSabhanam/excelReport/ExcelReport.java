package com.ramSabhanam.excelReport;

import java.util.List;
import java.util.Map;

import com.ramSabhanam.Database.IDatabase;
import com.ramSabhanam.Database.MicrosoftExcel;

public class ExcelReport {

	IDatabase database = null;
	
	public static ExcelReport reporter() { return ExcelReport.localProps.get(); }

	
	private static InheritableThreadLocal<ExcelReport> localProps = new InheritableThreadLocal<ExcelReport>(){
		@Override
		protected ExcelReport initialValue(){
			ExcelReport excelReport = new ExcelReport();
			return excelReport;
		}
	};
	
	public void setDatabase(String databaseType, String databaseLocation) {
		
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
