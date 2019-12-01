package com.ramSabhanam.tests;

import static com.ramSabhanam.excelReport.GlobalData.data;

import org.testng.annotations.Test;

public class ExcelReportTests {

	@Test
	public void test01() {
		
		data().put("Column1", "asdas");
		data().put("Sairam", "asd");
		data().put("Thirumala", "asx");
		data().put("ExecutionStatus", "PASSED");
		data().put("qwsd", "sdwd");
		
	}
	
}
