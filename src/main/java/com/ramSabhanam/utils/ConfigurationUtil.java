package com.ramSabhanam.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

public class ConfigurationUtil {

	public static void main(String[] args) {
		
	}
	
	public static Map<String,String> getBundle() { return ConfigurationUtil.localProps.get(); }

	private static InheritableThreadLocal<Map<String,String>> localProps = new InheritableThreadLocal<Map<String,String>>(){
		@Override
		protected Map<String,String> initialValue(){
			Properties properties = new Properties();
			File file = new File("excelReport.properties");
			try {
				properties.load(new FileInputStream(file));
			} catch (IOException e) {
				e.printStackTrace();
			}
			
			Map<String,String> props = new HashMap<String, String>();
			properties.forEach((k,v) -> {
				props.put(k.toString(), v.toString());
			});
			return props;
		}
	};

}
