package com.ramSabhanam.utils;

import java.text.SimpleDateFormat;
import java.util.Date;

public class DateUtil {

	public String getCurrentDateInFormat(String format) {
		
		SimpleDateFormat sdf = new SimpleDateFormat(format);
		Date date = new Date();
		return sdf.format(date);
			
	}

}
