package com.sysev;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.sysev.model.RowData;



public class GroupDataService {
	private enum Format{
		INT,
		CURRENCY
	}
	
	public static final SimpleDateFormat MMDDYYYY = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
	public static final Integer PERIOD = 3;
	public static final Integer FISCAL_YEAR = 5;
	public static final String NEW_LINE = System.getProperty("line.separator");
	public static Map<Integer, Format> FORMAT_MAP;
	
	static{
		FORMAT_MAP = new HashMap<Integer, Format>();
		FORMAT_MAP.put(0, Format.INT);
		FORMAT_MAP.put(1, Format.INT);
		FORMAT_MAP.put(2, Format.INT);
		FORMAT_MAP.put(3, Format.INT);
		FORMAT_MAP.put(5, Format.INT);
		FORMAT_MAP.put(6, Format.CURRENCY);
		FORMAT_MAP.put(7, Format.CURRENCY);
		FORMAT_MAP.put(8, Format.CURRENCY);
		FORMAT_MAP.put(9, Format.CURRENCY);
		FORMAT_MAP.put(10, Format.CURRENCY);
		FORMAT_MAP.put(11, Format.CURRENCY);
		FORMAT_MAP.put(12, Format.CURRENCY);
		FORMAT_MAP.put(14, Format.INT);
		FORMAT_MAP.put(15, Format.CURRENCY);
		FORMAT_MAP.put(16, Format.INT);
		FORMAT_MAP.put(19, Format.INT);
		FORMAT_MAP.put(21, Format.INT);
		FORMAT_MAP.put(22, Format.INT);
		FORMAT_MAP.put(23, Format.INT);
		FORMAT_MAP.put(25, Format.INT);
		FORMAT_MAP.put(26, Format.INT);
		FORMAT_MAP.put(27, Format.INT);
		FORMAT_MAP.put(35, Format.INT);
		FORMAT_MAP.put(36, Format.INT);
		FORMAT_MAP.put(39, Format.INT);
		
	}
	
	
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		try{
			File sourceDir = new File("./");
			System.out.println("sourceDir = " + sourceDir.getAbsolutePath());
			File outputDir = new File(sourceDir.getAbsolutePath() + "\\output");
			FileInputStream fis = null;
			try {
				if (outputDir.exists()) {
					delete(outputDir);
				}
				outputDir.mkdir();
				File[] files = sourceDir.listFiles();
				int totalRows = 0;
				int totalFiles = 0;
				if (files != null) {
					StringBuilder sb = new StringBuilder();
					sb
						.append("====================================================")
						.append(NEW_LINE + "Processing started: " + MMDDYYYY.format(new Date()))
						.append(NEW_LINE + "====================================================");
					writeToLog(sb);
					for (int f = 0; f < files.length; f++) {
						sb = new StringBuilder();
						File file = files[f];
						if (file.isFile() && (file.getName().endsWith(".xls2") || file.getName().endsWith(".xlsx"))) {
							sb.append(NEW_LINE + "Processing "	+ file.getAbsolutePath());
							//
							// Create a FileInputStream that will be use to read the
							// excel file.
							//
							fis = new FileInputStream(file);
							//
							// Create an excel workbook from the file system.
							//
							HSSFWorkbook workbook = new HSSFWorkbook(fis);
							
							
							//
							// Get the sheets in the workbook
							//
							int n = workbook.getNumberOfSheets();
							int rowsInFile = 0;
							for (int s = 0; s < n; s++) {
								Sheet sheet = workbook.getSheetAt(s);
								sb.append(NEW_LINE + " Processing Sheet " + (s + 1));
								Iterator<Row> rows = sheet.rowIterator();
								int rowsInSheet = 0;
								while (rows.hasNext()) {
									Row row = rows.next();
									try {
										RowData rowData2 = new RowData();
										String[] rowData = rowData2.getData();
										Cell cell = row.getCell(0);
										if (rowsInSheet == 0 && cell.getCellType() == HSSFCell.CELL_TYPE_STRING && "PARENT_COMPANY_ID".equalsIgnoreCase(cell.getStringCellValue())) {
											if (rows.hasNext()) {
												row = (HSSFRow)rows.next();
												cell = row.getCell(0);
											}
										}
										rowData[0] = getValue(cell, 0);
										for(int j = 1; j <=56; j++){
											cell = row.getCell(j);
											rowData[j] = getValue(cell, j);
										}
										String outputFileName = outputDir
												.getAbsolutePath()
												+ "\\"
												+ rowData[FISCAL_YEAR]
												+ new DecimalFormat("00").format((Double.valueOf(rowData[PERIOD]) - 3.0))
												+ ".txt";
										System.out.println(outputFileName);
										File output = new File(outputFileName);
										boolean newline = !output.createNewFile();
										FileWriter fileWriter = null;
										BufferedWriter bufferedWriter = null;
										try {
											fileWriter = new FileWriter(output, true);
											bufferedWriter = new BufferedWriter(fileWriter);
											if (newline) {
												bufferedWriter.newLine();
											}
											bufferedWriter.write(rowData2.toString());
											//System.out.println(rowData2);
										} catch (IOException ex) {
											ex.printStackTrace();
										} finally {
											try {
												bufferedWriter.close();
												fileWriter.close();
												bufferedWriter = null;
												fileWriter = null;
											} catch (Exception ex) {
												ex.printStackTrace();
											}
										}
									} catch (NoSuchElementException nsee) {
										// ok
									}
									rowsInSheet++;
								}
								sb.append(NEW_LINE + "Rows processed in sheet " + (s + 1) + ": " + rowsInSheet);
								rowsInFile += rowsInSheet;
							}
							sb
								.append(NEW_LINE + "-----------------------------")
								.append(NEW_LINE + "Total Rows processed in file \"" + file.getAbsolutePath() + "\": " + rowsInFile)
								.append(NEW_LINE + NEW_LINE);
							writeToLog(sb);
							sb = null;
							totalRows += rowsInFile;
							totalFiles++;
							fis.close();
						}
					}
				}
				
				StringBuilder sb = new StringBuilder();
				sb
					.append(NEW_LINE)
					.append(NEW_LINE + "--- SUMMARY ---")
					.append(NEW_LINE + "Total number of files processed: " + totalFiles)
					.append(NEW_LINE + "Total rows processed: " + totalRows)
					.append(NEW_LINE + NEW_LINE + "====================================================")
					.append(NEW_LINE + "Processing finished: " + MMDDYYYY.format(new Date()))
					.append(NEW_LINE + "====================================================");
				
				writeToLog(sb);
				sb = null;
			} catch (IOException e) {
				e.printStackTrace();
			} finally {
				if (fis != null) {
					try {
						fis.close();
					} catch (IOException ioe) {
	
					}
				}
			}
		}catch(Exception e){
			e.printStackTrace();
		}
	}

	public static void writeToLog(StringBuilder sb){
		FileWriter fileWriter = null;
		File log = new File("./output/log.txt");
		try {
			log.createNewFile();
			fileWriter = new FileWriter(log, true);
			fileWriter.write(sb.toString());
		} catch (IOException ex) {
			ex.printStackTrace();
		} finally {
			try {
				fileWriter.close();
			} catch (Exception ex) {
				ex.printStackTrace();
			}
		}
		sb = null;
	}
	
	public static String getValue(Cell cell, int index){
		if(cell == null) return "";
		switch(cell.getCellType()){
			case Cell.CELL_TYPE_STRING:
				return cell.getStringCellValue();
			case Cell.CELL_TYPE_NUMERIC:
				Double value = cell.getNumericCellValue();
				NumberFormat nf = null;
				Format f = FORMAT_MAP.get(index);
				if(f != null){
					switch(f){
						case INT:
							nf = new DecimalFormat("#");
							break;
						case CURRENCY:
							nf = new DecimalFormat("0.00");
							break;
					}
				}
				if(nf == null){
					return String.valueOf(value);
				}else{
					return nf.format(value);
				}
		}
		return "";
	}
	
	public static void delete(File file) throws IOException {

		if (file.isDirectory()) {

			// directory is empty, then delete it
			if (file.list().length == 0) {

				file.delete();
				System.out.println("Directory is deleted : "
						+ file.getAbsolutePath());

			} else {

				// list all the directory contents
				String files[] = file.list();

				for (String temp : files) {
					// construct the file structure
					File fileDelete = new File(file, temp);

					// recursive delete
					delete(fileDelete);
				}

				// check the directory again, if empty then delete it
				if (file.list().length == 0) {
					file.delete();
					System.out.println("Directory is deleted : "
							+ file.getAbsolutePath());
				}
			}

		} else {
			// if file, then delete it
			file.delete();
			System.out.println("File is deleted : " + file.getAbsolutePath());
		}
	}
}