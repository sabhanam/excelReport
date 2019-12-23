package com.ramSabhanam.tests;

import static com.ramSabhanam.excelReport.GlobalData.data;

import org.testng.annotations.Test;

public class ExcelReportTests {

	@Test
	public void test01() {
		
		data().put("Column1", "data1");
		data().put("Column2", "data2");
		data().put("Column3", "data3");
		data().put("Name", "sree");
		data().put("Address", "Street1");
		
	}
	
}
