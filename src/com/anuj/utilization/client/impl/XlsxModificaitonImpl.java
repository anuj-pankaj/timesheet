package com.anuj.utilization.client.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.anuj.utilization.client.inf.XlsxModificaiton;
import com.anuj.utilization.constant.Constants;

public class XlsxModificaitonImpl implements XlsxModificaiton {
	
	private int lastIndexColumnAddEffort=0;

	@Override
	public void addEeffortsCoumnn(String mainFile) throws IOException {
		File file = getFile(mainFile);
		try (FileInputStream fileInputStream = new FileInputStream(file)) {

			try (Workbook workbook = new XSSFWorkbook(fileInputStream)) {
				Sheet sheet = workbook.getSheetAt(0);
				System.out.println("sheet name : " + sheet.getSheetName());
				Font font =workbook.createFont();
				font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
				 CellStyle cellStyle = workbook.createCellStyle();
				cellStyle.setFont(font);
				Row row = sheet.getRow(0);

//				int lastCellNum = row.getLastCellNum();
				if(lastIndexColumnAddEffort==0) {
					lastIndexColumnAddEffort = row.getLastCellNum();
				}
				Cell existingLastCell = row.getCell(lastIndexColumnAddEffort-1);
				if (!existingLastCell.getStringCellValue().contains("Actual Time spent"))
				{
					Cell secondLastCell = row.createCell(lastIndexColumnAddEffort, CellType.STRING);
					secondLastCell.setCellStyle(cellStyle);
					secondLastCell.setCellValue(OFFICE_HOURS);

					Cell lastCell = row.createCell(lastIndexColumnAddEffort + 1, CellType.STRING);
					lastCell.setCellStyle(cellStyle);
					lastCell.setCellValue(OUT_OFFICE_HOURS);

				}
				try(FileOutputStream fileOutputStream = new FileOutputStream(file)) {
					workbook.write(fileOutputStream);
				}
			}
		}
	}

	@Override
	public void updateEfforts(String directory,List<String> fileNames, String startDate, String endDate) throws IOException, ParseException {
		File file  = getFile(getMainFile(directory));
		try (FileInputStream fileInputStream = new FileInputStream(file)) {

			try (Workbook workbook = new XSSFWorkbook(fileInputStream)) {
				Sheet sheet = workbook.getSheetAt(0);
				System.out.println("sheet name : " + sheet.getSheetName());
				
				List<Row> acitivityRows = getRows(fileNames,directory, 0, 2, startDate, endDate);
				Map<String, Map<String,Double>>  assignedIncidents = getTaskEffort(acitivityRows);
				Iterator<Row> iterator= sheet.iterator();
				Set<String> incidents = assignedIncidents.keySet();
				while(iterator.hasNext()) {
					Row row = iterator.next();
					if(row.getCell(0)!=null && incidents.contains(row.getCell(0).getStringCellValue()))
					{
						Map<String, Double> incidentValue = assignedIncidents.get(row.getCell(0).getStringCellValue());
						System.out.println("incident : " + row.getCell(0).getStringCellValue());
						
//						int lastRowIndex = row.getLastCellNum();
						if(incidentValue.get(IN_OFFICE_EFFORT)!=null ) {
							Cell inOfficeEffort = row.createCell(lastIndexColumnAddEffort, CellType.NUMERIC);
							inOfficeEffort.setCellValue(incidentValue.get(IN_OFFICE_EFFORT));
							System.out.println( " last cell value : " + incidentValue.get(IN_OFFICE_EFFORT));
						}
						
						if(incidentValue.get(OUT_OFFICE_EFFORT)!=null ) {
							Cell outOfficeEffort = row.createCell(lastIndexColumnAddEffort+1, CellType.NUMERIC);
							outOfficeEffort.setCellValue(incidentValue.get(OUT_OFFICE_EFFORT));
							System.out.println( " second last cell : " + incidentValue.get(OUT_OFFICE_EFFORT));
						}
							
					}
				}
				
				try(FileOutputStream fileOutputStream = new FileOutputStream(file)) {
					workbook.write(fileOutputStream);
				}
			}
		}
	}
		private Map<String, Map<String, Double>> getTaskEffort(List<Row> activityRows) {
		Map<String, Map<String, Double>> map = new HashMap<>();
		for (Row row : activityRows) {
		
				Cell celldate = row.getCell(2);
				if (celldate != null) {
					Map<String, Double> mapValue = new HashMap<>();
					Cell cellEffortsInOffice = row.getCell(4);
					if (cellEffortsInOffice != null) {
						double effortInOffice = cellEffortsInOffice.getNumericCellValue();
						
						mapValue.put(IN_OFFICE_EFFORT, effortInOffice);
						map.put(row.getCell(0).getStringCellValue().trim(), mapValue);
					} 
						Cell cellEffortsOutOffice = row.getCell(5);
						if (cellEffortsOutOffice != null) {
							double effortOutOffice = cellEffortsOutOffice.getNumericCellValue();
							mapValue.put(OUT_OFFICE_EFFORT, effortOutOffice);
							
						} 
						map.put(row.getCell(0).getStringCellValue().trim(), mapValue);
				}

		}
		return map;
	}


	private List<Row> getRows(List<String> fileNames, String directory, int sheetNumber, int cellIndexToEstimate,
			String startDate, String endDate) throws IOException {

		List<Row> rows = new ArrayList<Row>();
		for (String file : fileNames) {
			String fileName = directory + "\\" + file;
			try (FileInputStream fileInputStream = new FileInputStream(new File(fileName))) {
				System.out.println("File Name : " + fileName);

				try (Workbook workbook = new XSSFWorkbook(fileInputStream)) {
					Sheet sheet = workbook.getSheetAt(sheetNumber);
					System.out.println("sheet name : " + sheet.getSheetName());
					Iterator<Row> iterator = sheet.iterator();
					int count = 0;
					while (iterator.hasNext()) {

						Row row = iterator.next();
						if (count > 0) {
							Cell celldate = row.getCell(cellIndexToEstimate);
							if (celldate != null) {
								String workedDate = getDate(celldate.getNumericCellValue());
								System.out.println("Incident Number : " + row.getCell(0).getStringCellValue() + " -> "
										+ " worked date : " + workedDate + " -> " + "status : "
										+ row.getCell(3).getStringCellValue());
								if (isWorkDateInRange(workedDate, startDate, endDate)
										&& "Resolved".equalsIgnoreCase(row.getCell(3).getStringCellValue())) {
									Cell cellEffortsInOffice = row.getCell(4);
									if("INC001186532".equalsIgnoreCase(row.getCell(0).getStringCellValue()))
									{
										System.out.println();
									}
									if (cellEffortsInOffice != null) {
										rows.add(row);
									}
									Cell cellEffortsOutOffice = row.getCell(5);
									if (cellEffortsOutOffice != null) {
										rows.add(row);
									}
								}
							}

						}
						count++;

					}
				}
			}
		}
		return rows;
	}
	
	
	private File getFile(String filePath) {
		return new File(filePath);
	}
	
	private String getDate(double dateInDouble) {
		Date date = DateUtil.getJavaDate(dateInDouble);
		return new SimpleDateFormat("dd-MM-yyyy").format(date);
	}
	
	private boolean isWorkDateInRange(String workedDate, String startDate, String endDate) {
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("[yyyy-MM-dd][dd/MM/yyyy][dd-MM-yyyy]");
		LocalDate  workedDateLocalDate = LocalDate.parse(workedDate,formatter);
		if(workedDateLocalDate.isAfter(LocalDate.parse(startDate,formatter).minusDays(1))&& workedDateLocalDate.isBefore(LocalDate.parse(endDate,formatter).plusDays(1))) {
			return true;
		}
		return false;
	}
	

	@Override
	public void mergeSheets(List<String> fileNames, String directory, String startDate, String endDate)
			throws IOException {
		// Merge Non Assigned tickets
		mergeSheet(fileNames, directory,Constants.SHEET_TWO_HEADER ,"Non assinged tickets", 1, 34, startDate, endDate);
		// Merge others
		mergeSheet(fileNames, directory,Constants.SHEET_THREE_HEADER, "Others", 2, 3, startDate, endDate);
		// Merge CTASK
		mergeSheet(fileNames, directory,Constants.SHEET_FOUR_HEADER, "CTASK", 3, 5, startDate, endDate);
	}


	
	public void mergeSheet(List<String> fileNames, String directory,String[] headerColumns, String newSheetName,int sheetIndex,int workedDateIndex , String startDate, String endDate)
			 throws IOException{
		try (FileInputStream mainFileInputStream = new FileInputStream(getFile(getMainFile(directory)))) {
			try (Workbook workbookOutPut = new XSSFWorkbook(mainFileInputStream)) {
				Sheet sheetOutPut = workbookOutPut.createSheet(newSheetName);
				
				Row headerRow = sheetOutPut.createRow(0);
				prepareHeader(headerRow,workbookOutPut.createFont(),workbookOutPut.createCellStyle(),headerColumns);
				int rowIndexes = 1;
				for (String file : fileNames) {
					String fileName = directory + file;
					try (FileInputStream fileInputStream = new FileInputStream(new File(fileName))) {
						try (Workbook workbookInput = new XSSFWorkbook(fileInputStream)) {
							Sheet sheetInput = workbookInput.getSheetAt(sheetIndex);
							
							Iterator<Row> iteratorRow = sheetInput.iterator();
							int skipFirstRow = 0;
							while (iteratorRow.hasNext()) {
								Row rowInput = iteratorRow.next();
								if (skipFirstRow > 0 && rowInput.getCell(0)!=null) {
									String workedDate = null;
									try {
										System.out.println("First cell data " + rowInput.getCell(0).getStringCellValue());
										workedDate = getDate(rowInput.getCell(workedDateIndex).getNumericCellValue());
									}
									catch (IllegalStateException e) {
										System.out.println();
									}
								//	String workedDate = getDate(rowInput.getCell(workedDateIndex).getNumericCellValue());
									if (isWorkDateInRange(workedDate, startDate, endDate)) {
										int lastCellIndex = rowInput.getLastCellNum();
										Row rowDestination = sheetOutPut.createRow(rowIndexes);
										rowIndexes++;
										for (int i = 0; i < lastCellIndex; i++) {
											Cell sourceCell = rowInput.getCell(i);
											Cell destinationCell = rowDestination.createCell(i);
											setCellvalue(sourceCell, destinationCell);
										}
									}

								}
								skipFirstRow++;
							}

						}
					}
					
				}
				try (FileOutputStream fileOutputStream = new FileOutputStream(getFile(getMainFile(directory)))) {
					workbookOutPut.write(fileOutputStream);
				}

			}
		}

	}
	
	private String getMainFile(String directory) {
		return directory+"incident.xlsx";
	}
	private void prepareHeader(Row headerRow, Font font,CellStyle cellStyle , String[] headerColumns) {
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		cellStyle.setFont(font);
		for(int i=0;i<headerColumns.length;i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(headerColumns[i]);
		}
		
	}
	
	
	private void setCellvalue(Cell sourceCell,Cell destinationCell) {
		if(sourceCell!=null) {
			destinationCell.setCellType(sourceCell.getCellType());
	        switch (sourceCell.getCellType()) {
	            case Cell.CELL_TYPE_BLANK:
	            	destinationCell.setCellValue(sourceCell.getStringCellValue());
	                break;
	            case Cell.CELL_TYPE_BOOLEAN:
	            	destinationCell.setCellValue(sourceCell.getBooleanCellValue());
	                break;
	            case Cell.CELL_TYPE_ERROR:
	            	destinationCell.setCellErrorValue(sourceCell.getErrorCellValue());
	                break;
	            case Cell.CELL_TYPE_FORMULA:
	            	destinationCell.setCellFormula(sourceCell.getCellFormula());
	                break;
	            case Cell.CELL_TYPE_NUMERIC:
	            	destinationCell.setCellValue(sourceCell.getNumericCellValue());
	                break;
	            case Cell.CELL_TYPE_STRING:
	            	destinationCell.setCellValue(sourceCell.getRichStringCellValue());
	                break;
	        }
		}else {
			destinationCell.setCellValue("");
		}
		
	}
}
