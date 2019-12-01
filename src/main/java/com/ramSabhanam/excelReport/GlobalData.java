package com.ramSabhanam.excelReport;

import java.util.HashMap;
import java.util.Map;

public class GlobalData {

	public static Map<String,String> data() { return GlobalData.localProps.get(); }

	
	private static InheritableThreadLocal<Map<String,String>> localProps = new InheritableThreadLocal<Map<String,String>>(){
		@Override
		protected Map<String,String>initialValue(){
			Map<String,String> data = new HashMap<>();
			return data;
		}
	};
}
