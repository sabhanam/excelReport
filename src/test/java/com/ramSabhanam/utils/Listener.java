package com.ramSabhanam.utils;

import static com.ramSabhanam.configuration.Configuration.*;
import static com.ramSabhanam.excelReport.ExcelReport.reporter;
import static com.ramSabhanam.excelReport.GlobalData.data;

import org.testng.ITestContext;
import org.testng.ITestListener;
import org.testng.ITestResult;

public class Listener implements ITestListener{

	@Override
	public void onTestStart(ITestResult result) {
		/**
		 * #############################################
		 * 
		 * On start of each test, adding Test Case Name
		 * 
		 * #############################################
		 */
		data().put("TestCaseName", result.getName());
	}

	@Override
	public void onTestSuccess(ITestResult result) {
		/**
		 * #############################################
		 * 
		 * Call this flushExcel with data() from GlobalData, see README.md for usage
		 * After flushing, clear the data()
		 * 
		 * #############################################
		 */
		data().put("ExecutionStatus", "PASSED");
		reporter().flushExcel(data());
		data().clear();
	}

	@Override
	public void onTestFailure(ITestResult result) {
		/**
		 * #############################################
		 * 
		 * Call this flushExcel with data() from GlobalData, see README.md for usage
		 * After flushing, clear the data()
		 * 
		 * #############################################
		 */
		data().put("ExecutionStatus", "FAILED");
		reporter().flushExcel(data());
		data().clear();
	}

	@Override
	public void onTestSkipped(ITestResult result) {

	}

	@Override
	public void onTestFailedButWithinSuccessPercentage(ITestResult result) {

	}

	@Override
	public void onStart(ITestContext context) {
		/**
		 * #############################################
		 * 
		 * Initiate the Report Name
		 * 
		 * #############################################
		 */
		reporter().setDatabase("xlsx", getBundle().get("reports.location") + "/" + context.getName());
	}

	@Override
	public void onFinish(ITestContext context) {
		/**
		 * #############################################
		 * 
		 * On complete
		 * 
		 * #############################################
		 */
		reporter().dispose();
	}

}
