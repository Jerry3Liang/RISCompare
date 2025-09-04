package com.uitc;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class DatetimeConverter {
	public static final SimpleDateFormat logTimeFormatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	public static final SimpleDateFormat logTimeFormatter2 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");
	public static final SimpleDateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd");
	public static final SimpleDateFormat dateFormatter2 = new SimpleDateFormat("yyyy年MM月dd日");
	public static final SimpleDateFormat dateFormatter3 = new SimpleDateFormat("yyyy/M/d");
	public static final SimpleDateFormat dateFormatter4 = new SimpleDateFormat("yyyyMMdd");

	public static String toString(Date datetime, String format) {
		String result = "";
		try {
			if (datetime != null) {
				result = new SimpleDateFormat(format).format(datetime);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}

	public static Date parse(String datetime, String format) {

		Date result;
		try {
			result = new SimpleDateFormat(format).parse(datetime);
		} catch (Exception e) {
			result = new Date();
			e.printStackTrace();
		}
		return result;
	}

	/*-
	 * get systime time
	 *
	 * @param type
	 *   0 : hhmmss
	 *   1 : long time
	 *   2 : yyyy-MM-dd HH:mm:ss
	 *   3 : yyyy-MM-dd
	 *   4 : yyyy年MM月dd日
	 *   5 : yyyy/M/d
	 *   6 : yyyyMMdd
	 *   7 : yyyy-MM-dd HH:mm:ss.SSS
	 * @return : String
	 */
	public static String getSYSTime(int type) {

		String systime = "";
		Calendar now = Calendar.getInstance();

		switch (type) {
			case 0:
				systime = logTimeFormatter.format(now.getTime()).substring(8, 14);
				break;
			case 1:
				long endtime = (new Date()).getTime();
				systime = String.valueOf(endtime);
				break;
			case 2:
				systime = logTimeFormatter.format(now.getTime());
				break;
			case 3:
				systime = dateFormatter.format(now.getTime());
				break;
			case 4:
				systime = dateFormatter2.format(now.getTime());
				break;
			case 5:
				systime = dateFormatter3.format(now.getTime());
				break;
			case 6:
				systime = dateFormatter4.format(now.getTime());
				break;
			case 7:
				systime = logTimeFormatter2.format(now.getTime());
				break;
			default:
				break;
		}

		return systime;
	}

}
