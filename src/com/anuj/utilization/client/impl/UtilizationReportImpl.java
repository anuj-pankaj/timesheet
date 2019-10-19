package com.anuj.utilization.client.impl;

import java.io.File;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import com.anuj.utilization.client.inf.MergingFilesService;
import com.anuj.utilization.client.inf.UtilizationReport;
import com.anuj.utilization.client.inf.XlsxModificaiton;


public class UtilizationReportImpl implements UtilizationReport {

	private String directory;
	private String startDate;
	private String endDate;
	private List<String> mergingFiles ;
	MergingFilesService  mergingFileService;
	XlsxModificaiton xlsxModificaiton;
	
	public UtilizationReportImpl(String directory,String startDate, String endDate) {
		super();
		this.directory = directory;
		this.startDate = startDate;
		this.endDate = endDate;
		xlsxModificaiton = new XlsxModificaitonImpl();
	}

	@Override
	public void prepareUtilizatoinReport() {
		try {
			setMergingFiles(getDirectory());
			List<String> fileNames = getMergingFiles();
			fileNames.stream().forEach(fileName->System.out.println(fileName));
			String absoluteFilePath = directory+"incident.xlsx";
			xlsxModificaiton.addEeffortsCoumnn(absoluteFilePath);
			xlsxModificaiton.updateEfforts(getDirectory(),fileNames,getStartDate(),getEndDate());
			xlsxModificaiton.mergeSheets(fileNames, getDirectory(), getStartDate(),getEndDate());
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("Error message : " + e.getMessage());
		}
			
		mergingFileService = new MergingFilesServiceImpl();
		mergingFileService.mergeIncidentSheet(directory);
	}



	public List<String> getMergingFiles() {
		return mergingFiles;
	}


	public String getDirectory() {
		return directory;
	}

	public String getStartDate() {
		return startDate;
	}

	public String getEndDate() {
		return endDate;
	}

	private void setMergingFiles(String filePath ) throws Exception {
	
		
		File file = new File(filePath);
		if(file.isDirectory()) {
			String [] fileNames = file.list();
			mergingFiles = Arrays.asList(fileNames).stream().filter(fileName->  !"incident.xlsx".equalsIgnoreCase(fileName)).collect(Collectors.toList()); 
		} else {
			throw new Exception("This is not a  File Directory");
		}
	
	}

}
