package com.sysev;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.RowRecord;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DateUtil;

import com.sysev.model.RowData;
import com.sysev.util.FileUtil;

public class GroupDataService2 implements HSSFListener{
	private File outputDir = null;
	
	public static final String NEW_LINE = System.getProperty("line.separator");
	public static Map<Short, NumberFormat> FORMAT_MAP;
	public static final Map<Integer, String> MONTH_NAMES;
	
	static{
		NumberFormat intFormat = new DecimalFormat("#");
		NumberFormat curFormat = new DecimalFormat("0.00");
		
		FORMAT_MAP = new HashMap<Short, NumberFormat>();
		FORMAT_MAP.put((short)0,intFormat);
		FORMAT_MAP.put((short)1,intFormat);
		FORMAT_MAP.put((short)2,intFormat);
		FORMAT_MAP.put((short)3,intFormat);
		FORMAT_MAP.put((short)5,intFormat);
		FORMAT_MAP.put((short)6,curFormat);
		FORMAT_MAP.put((short)7,curFormat);
		FORMAT_MAP.put((short)8,curFormat);
		FORMAT_MAP.put((short)9,curFormat);
		FORMAT_MAP.put((short)10,curFormat);
		FORMAT_MAP.put((short)11,curFormat);
		FORMAT_MAP.put((short)12,curFormat);
		FORMAT_MAP.put((short)14,intFormat);
		FORMAT_MAP.put((short)15,curFormat);
		FORMAT_MAP.put((short)16,intFormat);
		FORMAT_MAP.put((short)19,intFormat);
		FORMAT_MAP.put((short)21,intFormat);
		FORMAT_MAP.put((short)22,intFormat);
		FORMAT_MAP.put((short)23,intFormat);
		FORMAT_MAP.put((short)25,intFormat);
		FORMAT_MAP.put((short)26,intFormat);
		FORMAT_MAP.put((short)27,intFormat);
		FORMAT_MAP.put((short)35,intFormat);
		FORMAT_MAP.put((short)36,intFormat);
		FORMAT_MAP.put((short)39,intFormat);
		
		MONTH_NAMES = new HashMap<Integer, String>();
		MONTH_NAMES.put(0, "January");
		MONTH_NAMES.put(1, "February");
		MONTH_NAMES.put(2, "March");
		MONTH_NAMES.put(3, "April");
		MONTH_NAMES.put(4, "May");
		MONTH_NAMES.put(5, "June");
		MONTH_NAMES.put(6, "July");
		MONTH_NAMES.put(7, "August");
		MONTH_NAMES.put(8, "September");
		MONTH_NAMES.put(9, "October");
		MONTH_NAMES.put(10, "November");
		MONTH_NAMES.put(11, "December");
	}	
    private SSTRecord sstrec;
    private int currentRow = -1;
    private int currentSheet = -1;
    private boolean currentRowIsHeader = false;
    private long rowCount = 0;
    private RowData currentRowData = null;
    
    public long getRowCount() {
		return rowCount;
	}

	public void setRowCount(long rowCount) {
		this.rowCount = rowCount;
	}

	public GroupDataService2(File outputDir){
		this.outputDir = outputDir;
	}
	
	/**
     * This method listens for incoming records and handles them as required.
     * @param record    The record that was found while reading.
     */
    public void processRecord(Record record)
    {
        switch (record.getSid())
        {
            // the BOFRecord can represent either the beginning of a sheet or the workbook
            case BOFRecord.sid:
                BOFRecord bof = (BOFRecord) record;
                if (bof.getType() == BOFRecord.TYPE_WORKBOOK)
                {
                    System.out.println("Encountered workbook");
                    currentSheet = -1;
                    // assigned to the class level member
                } else if (bof.getType() == BOFRecord.TYPE_WORKSHEET) {
                    System.out.println("Encountered sheet reference");
                    writeToLog(NEW_LINE + "Sheet " + (++currentSheet + 1));
                }
                break;
            case BoundSheetRecord.sid:
                BoundSheetRecord bsr = (BoundSheetRecord) record;
                System.out.println("Bound sheet named: " + bsr.getSheetname());
                break;
            case RowRecord.sid:
                //RowRecord row = (RowRecord) record;
                rowCount++;
                break;
            case SSTRecord.sid:
                sstrec = (SSTRecord) record;
                break;                
            case NumberRecord.sid:
            case LabelSSTRecord.sid:
            	process(record);
                break;
        }
    }

    private void process(Record cell){
    	String value = "";
    	int row = -1;
    	short col = -1;
    	if(cell instanceof NumberRecord){
    		NumberRecord ncell = (NumberRecord)cell;
    		row = ncell.getRow();
    		col = ncell.getColumn();
    		if(col == RowData.AGING_DATE || col == RowData.POSTING_DATE || col == RowData.DOB || col == RowData.SERVICE_MONTH){
    			double d = ncell.getValue();
    			if(d != 0){
    				try{
    					Date date = DateUtil.getJavaDate(d);
    					value = RowData.DDMMMYYYY.format(date);
    				}catch(NullPointerException npe){
    					writeToLog(NEW_LINE + " * Invalid date in cell " + getExcelColumnName(col + 1) + (row + 1) + ": " + ncell.getValue());
    				}
    			}
    		}else if(col == RowData.MONTH){
    			double d = ncell.getValue();
    			if(d != 0){
    				try{
    					Date date = DateUtil.getJavaDate(d);
    					value = RowData.DDMMMYYYY.format(date);
    				}catch(NullPointerException npe){
    					writeToLog(NEW_LINE + " * Invalid date in cell " + getExcelColumnName(col + 1) + (row + 1) + ": " + ncell.getValue());
    				}   				
    			}
    		}else{
	    		NumberFormat f = FORMAT_MAP.get(col);
	    		if(f == null){
	    			value = String.valueOf(ncell.getValue());
	    		}else{
	    			try{
	    				value = f.format(ncell.getValue());
	    			}catch(ArithmeticException ae){
	    				writeToLog(NEW_LINE + " * Formatting error in cell " + getExcelColumnName(col + 1) + (row + 1) + ": " + ncell.getValue());
	    			}
	    		}
    		}
    	}else if(cell instanceof LabelSSTRecord){
			LabelSSTRecord lcell = (LabelSSTRecord)cell;
			row = lcell.getRow();
			col = lcell.getColumn();
			value = sstrec.getString(lcell.getSSTIndex()).toString();
		}else{
			System.out.println("UNKNOWN TYPE FOR: " + cell.toString());
		}
    	
		if(currentRow != row){
			currentRow = row;
			currentRowIsHeader = false;
		 	if(currentRowData != null){
		 		write(currentRowData);
		 	}
		 	currentRowData = new RowData();
		}
		 
		if(row == 0 && col == 0 && "PARENT_COMPANY_ID".equalsIgnoreCase(value)){
			currentRowIsHeader = true;
			rowCount--;
		}
		if(!currentRowIsHeader){
			String[] data = currentRowData.getData();
			//System.out.println("[" + currentRow + "," + col +"] " + value);
			data[col] = value;
		}
    }
    
    /**
     *
     * @throws IOException  When there is an error processing the file.
     */
    public static void main(String[] args) throws IOException
    {
    	File sourceDir = new File("./");
		System.out.println("sourceDir = " + sourceDir.getAbsolutePath());
		File outputDir = new File(sourceDir.getAbsolutePath() + "\\output");
		long totalRows = 0;
		int totalFiles = 0;
		try {
			if (outputDir.exists()) {
				FileUtil.delete(outputDir);
			}
			outputDir.mkdir();
			File[] files = sourceDir.listFiles();
			
			if (files != null) {
				StringBuilder sb = new StringBuilder();
				sb
					.append("====================================================")
					.append(NEW_LINE + "Processing started: " + RowData.MMDDYYYYHHmmss.format(new Date()))
					.append(NEW_LINE + "====================================================");
				sb = writeToLog(sb);
				
				for (int f = 0; f < files.length; f++) {
					sb = new StringBuilder();
					File file = files[f];
					if (file.isFile() && (file.getName().endsWith(".xls") || file.getName().endsWith(".xlsx"))) {
						totalFiles++;
						sb.append(NEW_LINE + "Processing "	+ file.getAbsolutePath());
						sb = writeToLog(sb);
						
						FileInputStream fis = null;
						InputStream din = null;
						try {
							fis = new FileInputStream(file);
					        POIFSFileSystem poifs = new POIFSFileSystem(fis);
					        try{
					        	din = poifs.createDocumentInputStream("Workbook");
					        }catch(FileNotFoundException fnfe){
					        	writeToLog(new StringBuilder("Workbook not found... trying \"Book\""));
								din = poifs.createDocumentInputStream("Book");
					        }
					        HSSFRequest req = new HSSFRequest();
					        GroupDataService2 service = new GroupDataService2(outputDir);
					        req.addListenerForAllRecords(service);
					        HSSFEventFactory factory = new HSSFEventFactory();
							
					        factory.processEvents(req, din);
					        fis.close();
					        din.close();
					        System.out.println(service.getRowCount());
					        totalRows += service.getRowCount();
						} catch (IOException ioe) {
							// TODO: handle exception
						} finally{
							try{
								if(fis != null){
									fis.close();
								}
								if(din != null){
									din.close();
								}
							}catch(Exception e){
								e.printStackTrace();
								writeExceptionToLog(e);
							}
						}
					}
				}
				sb
					.append(NEW_LINE)
					.append(NEW_LINE + "--- SUMMARY ---")
					.append(NEW_LINE + "Total number of files processed: " + totalFiles)
					.append(NEW_LINE + "Total rows processed: " + totalRows)
					.append(NEW_LINE + NEW_LINE + "====================================================")
					.append(NEW_LINE + "Processing finished: " + RowData.MMDDYYYYHHmmss.format(new Date()))
					.append(NEW_LINE + "====================================================");
				sb = writeToLog(sb);				
			}
		}catch(Exception e){
			e.printStackTrace();
			writeExceptionToLog(e);
		}
    }
		
    public void write(RowData row){
    	if(row == null) return;
    	String[] data = row.getData();
    	if(data.length < 5) return;
    	FileWriter fileWriter = null;
		BufferedWriter bufferedWriter = null;
    	try{
    		String strPeriod = data[RowData.PERIOD];
    		double period;
    		if(strPeriod != null){
    			try{
    				period = Double.parseDouble(strPeriod);
    				String outputFileName = outputDir.getAbsolutePath()
    				+ "\\"
    				+ data[RowData.FISCAL_YEAR]
    				+ new DecimalFormat("00").format(period - 3.0)
    				+ ".txt";
    				File output = new File(outputFileName);
    				boolean newline = !output.createNewFile();
    				fileWriter = new FileWriter(output, true);
    				bufferedWriter = new BufferedWriter(fileWriter);
    				if (newline) {
    					bufferedWriter.newLine();
    				}
    				bufferedWriter.write(row.toString());
    				
	    		}catch(NumberFormatException nfe){
	    			nfe.printStackTrace();
	    			writeExceptionToLog(nfe);
	    		}
    		}
		} catch (IOException ex) {
			ex.printStackTrace();
			writeExceptionToLog(ex);
		} finally {
			try {
				if(bufferedWriter != null){
					bufferedWriter.close();
					bufferedWriter = null;
				}
				if(fileWriter != null){
					fileWriter.close();
					fileWriter = null;
				}
			} catch (Exception ex) {
				ex.printStackTrace();
				writeExceptionToLog(ex);
			}
		}
    }
    
    private String getExcelColumnName(int columnNumber)
    {
        int dividend = columnNumber;
        String columnName = "";
        int modulo;

        while (dividend > 0)
        {
            modulo = (dividend - 1) % 26;
            columnName = new Character((char)(65 + modulo)) + columnName;
            dividend = (int)((dividend - modulo) / 26);
        } 

        return columnName;
    }
    
    public static void writeToLog(String str){
		FileWriter fileWriter = null;
		File log = new File("./output/log.txt");
		try {
			log.createNewFile();
			fileWriter = new FileWriter(log, true);
			fileWriter.write(str);
		} catch (IOException ex) {
			ex.printStackTrace();
		} finally {
			try {
				fileWriter.close();
			} catch (Exception ex) {
				ex.printStackTrace();
			}
		} 	
    }
    
	public static StringBuilder writeToLog(StringBuilder sb){
		writeToLog(sb.toString());
		return new StringBuilder();
	}   
	
	public static void writeExceptionToLog(Exception e){
		FileWriter fileWriter = null;
		File log = new File("./output/log.txt");
		try {
			log.createNewFile();
			fileWriter = new FileWriter(log, true);
			e.printStackTrace(new PrintWriter(fileWriter));
		} catch (IOException ex) {
			ex.printStackTrace();
		} finally {
			try {
				fileWriter.close();
			} catch (Exception ex) {
				ex.printStackTrace();
			}
		}
	} 	
}