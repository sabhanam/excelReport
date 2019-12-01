package com.ramSabhanam.Database;

import java.util.List;
import java.util.Map;

/**
 * Creation of Interface, in case of flushing this report utility into multiple DB's
 * 
 * @author RamSabhanam
 *
 */
public interface IDatabase {

	public void createDatabase(String databaseType, String databaseLocation);
	
	public boolean verifyDatabase(String databaseType, String databaseLocation);
	
	public void flushExcel(Map<String,String> data);
	
	public void flushExcel(List<Map<String,String>> data);
	
	public void dispose();
	
}
