package com.ramSabhanam.Database;

import static com.ramSabhanam.configuration.Configuration.getBundle;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ramSabhanam.date_util.*;

/**
 * Microsoft xlsx & xls related utility to flush report as Excel file
 * 
 * @author RamSabhanam
 *
 */
public class MicrosoftExcel implements IDatabase{

	private Workbook workbook = null;
	private String workbookLocation = "";
	Sheet sheet = null;
	
	/**
	 * 
	 * Creates the file, in case missing.
	 * Creation of default columns only happens first time.
	 * However, creation of columns which are added at runtime happens always.
	 * 
	 * @param databaseType : xls or xlsx
	 * @param databaseLocation : Location of file
	 * 
	 */
	public MicrosoftExcel(String databaseType, String databaseLocation) {
		
		workbookLocation = databaseLocation + "." + databaseType;
		
		if(!verifyDatabase(databaseType, databaseLocation)) {
			
			createDatabase(databaseType, databaseLocation);
			
		}
		
	}
	
	/**
	 * creation of Excel file
	 */
	@Override
	public void createDatabase(String databaseType, String databaseLocation) {
		
		if(databaseType.equals("xlsx")) {
			
			workbook = new XSSFWorkbook();
			
		}else {
			
			workbook = new HSSFWorkbook();
			
		}
		
	}

	/**
	 * Verification of Excel file, exists or not
	 */
	@Override
	public boolean verifyDatabase(String databaseType, String databaseLocation) {
		
		try {
			
			if(databaseType.equals("xlsx")) {
				
				workbook = new XSSFWorkbook(new FileInputStream(new File(databaseLocation + "." + databaseType)));
				
			}else {
				
				workbook = new HSSFWorkbook(new FileInputStream(new File(databaseLocation + "." + databaseType)));
				
			}
			
		}catch (Exception e) {
			
			return false;
			
		}
		
		return true;
		
	}
	
	/**
	 * Get columns which are default
	 * 
	 * @return ArrayList<String> : Column Names
	 */
	private ArrayList<String> getColumnNames() {
		
		ArrayList<String> columnNames = new ArrayList<String>();
		
		for(String columnName : getBundle().getOrDefault("report.default.heirarchy", "").split("\\,")) {
			columnNames.add(columnName);
		}
		
		return columnNames;
		
	}
	
	/**
	 * 
	 * Create the columns if they are unavailable
	 * 
	 * @param data
	 */
	private void makeAllColumnsAvailable(Map<String,String> data) {
		
		Row row = null;
		sheet = workbook.getSheet(getBundle().get("report.sheet.name"));
		
        int incrementor = 0;
        
        if(sheet == null) {
            
        	sheet = workbook.createSheet(getBundle().get("report.sheet.name"));
            row = sheet.createRow(0);
            
            if(!getBundle().getOrDefault("report.serialNumber.columnName", "").isEmpty()) {
            	incrementor = createColumnCell(row,incrementor,getBundle().getOrDefault("report.serialNumber.columnName", ""));
            }
            
            if(!getBundle().getOrDefault("report.timestamp.columnName", "").isEmpty()) {
            	incrementor = createColumnCell(row,incrementor,getBundle().getOrDefault("report.timestamp.columnName", ""));
            }
            
            if(!getBundle().getOrDefault("report.testName.columnName", "").isEmpty()) {
            	incrementor = createColumnCell(row,incrementor,getBundle().getOrDefault("report.testName.columnName", ""));
            }
            
            if(!getBundle().getOrDefault("report.executionStatus.columnName", "").isEmpty()) {
            	incrementor = createColumnCell(row,incrementor,getBundle().getOrDefault("report.executionStatus.columnName", ""));
            }
            
            for (String defaultColumn : getColumnNames()) {
            	incrementor = createColumnCell(row,incrementor,defaultColumn);
            }

        }
        
        Cell cell = null;
		
		row = sheet.getRow(0);
        incrementor = row.getLastCellNum();
        for (Map.Entry<String, String> entry : data.entrySet()) {
            boolean blnFlag = true;
            for (short i = 0; i < row.getLastCellNum() && blnFlag; i++) {
                cell = row.getCell(Integer.parseInt("" + i));
                if (cell.getStringCellValue().equals(entry.getKey())) {
                    blnFlag = false;
                }
            }
            if (blnFlag) {
            	incrementor = createColumnCell(row,incrementor,entry.getKey());
            }

        }
        
	}
	
	/**
	 * Creation of cell
	 * 
	 * @param row
	 * @param incrementor
	 * @param columnName
	 * @return
	 */
	private int createColumnCell(Row row, int incrementor, String columnName) {
		Cell cell = row.createCell(incrementor);
        cell.setCellStyle(getHeaderStyling());
        cell.setCellValue(columnName);
        incrementor++;
        return incrementor; 
	}

	/**
	 * Header cell styling
	 * 
	 * @return
	 */
	private CellStyle getHeaderStyling() {
		
		CellStyle style = workbook.createCellStyle();

        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);

        style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        
		return style;
		
	}
	
	/**
	 * Row cell styling
	 * 
	 * @return
	 */
	private CellStyle getRowStyling() {
		CellStyle style1 = workbook.createCellStyle();
		style1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		style1.setBorderTop(XSSFCellStyle.BORDER_THIN);
		style1.setBorderRight(XSSFCellStyle.BORDER_THIN);
		style1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		return style1;
	}
	
	/**
	 * Passed cell styling
	 * 
	 * @return
	 */
	private CellStyle getPassedCellStyling() {
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		return style;
	}
	
	/**
	 * Failed cell styling
	 * 
	 * @return
	 */
	private CellStyle getFailedCellStyling() {
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		return style;
	}

	/**
	 * Insert row into excel
	 * 
	 * @param data
	 */
	private void insertExecutionRow(Map<String, String> data) {
		
		ArrayList<String> columnNames = new ArrayList<>();
		
		sheet = workbook.getSheet(getBundle().get("report.sheet.name"));
		
		Row columnNamesRow = sheet.getRow(0);
		
		int rowNumber = sheet.getPhysicalNumberOfRows();
		
		int columnCount = sheet.getRow(0).getPhysicalNumberOfCells();
		
		for(int i = 0; i < columnCount; i++) {
			
			columnNames.add(columnNamesRow.getCell(i).getRichStringCellValue().toString());
			
		}
		
		Row appendingRow = sheet.createRow(rowNumber);
		DateUtil dateUtil = new DateUtil();
		
		for(int i = 0; i < columnCount; i++) {
			
			Cell cell = appendingRow.createCell(i);
			
			cell.setCellStyle(getRowStyling());
			
			if(columnNames.get(i).equals(getBundle().get("report.serialNumber.columnName"))) {
				
				cell.setCellValue(rowNumber);
				
			}else if(columnNames.get(i).equals(getBundle().get("report.timestamp.columnName"))) {
				
				cell.setCellValue(dateUtil.getCurrentDateInFormat("yyyy.MM.dd 'at' HH:mm:ss z"));
				
			}else if(columnNames.get(i).equals(getBundle().get("report.testName.columnName"))) {
				
				
				
			}else if(columnNames.get(i).equals(getBundle().get("report.executionStatus.columnName"))) {
				
				if(data.get(columnNames.get(i)).equals(getBundle().get("report.passed.text"))) {
					
					cell.setCellStyle(getPassedCellStyling());
					cell.setCellValue(getBundle().get("report.passed.text"));
					
				}else if(data.get(columnNames.get(i)).equals(getBundle().get("report.failed.text"))) {
					
					cell.setCellStyle(getFailedCellStyling());
					cell.setCellValue(getBundle().get("report.failed.text"));
					
				}
				
			}else {
				
				cell.setCellValue(data.get(columnNames.get(i)));
			
			}
			
		}
		
	}

	/**
	 * Insert data into excel
	 * 
	 */
	@Override
	public void flushExcel(Map<String, String> data) {
		
		makeAllColumnsAvailable(data);
		
		insertExecutionRow(data);
		
	}

	/**
	 * Insert multiple data into excel
	 */
	@Override
	public void flushExcel(List<Map<String, String>> data) {
		
		for(Map<String,String> eachData : data) {
			
			makeAllColumnsAvailable(eachData);
			
			insertExecutionRow(eachData);
			
		}
		
	}

	/**
	 * Save file at last
	 */
	@Override
	public void dispose() {
		
		try {
			
			FileOutputStream outFile =new FileOutputStream(workbookLocation);
	        workbook.write(outFile);
	        workbook.close();
	        
		}catch (Exception e) {
			
			e.printStackTrace();
			
		}
        
	}

}
