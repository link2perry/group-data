package com.sysev.model;

import java.text.SimpleDateFormat;


public class RowData {
	public static final SimpleDateFormat MMDDYYYYHHmmss = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
	public static final SimpleDateFormat DDMMMYYYY = new SimpleDateFormat("dd-MMM-yyyy");
	public static final Integer PERIOD = 3;
	public static final Integer MONTH = 4;
	public static final Integer FISCAL_YEAR = 5;
	public static final Integer SERVICE_MONTH = 24;
	public static final Integer AGING_DATE = 49;
	public static final Integer POSTING_DATE = 50;
	public static final Integer DOB = 51;	
	public static final String DELIMITER = "\t";
	
	String[] data = new String[57];
	
	public RowData(){}

	public String[] getData() {
		return data;
	}

	@Override
	public String toString() {
		StringBuilder sb = new StringBuilder(500);
		if(data == null){
			return "";
		}
		int n = data.length - 1;
		for(int i = 0; i < data.length; i++){
			if(data[i] == null){
				sb.append("");
			}else{
				sb.append(data[i]);
			}
			if(i < n){
				sb.append(DELIMITER);
			}
		}
		return sb.toString();
	}
	
	
}