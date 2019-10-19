package com.anuj.utilization.client.inf;

import java.io.IOException;
import java.text.ParseException;
import java.util.List;

public interface XlsxModificaiton {
	String OFFICE_HOURS="Actual Time spent on Incidence in UK Office (Minutes)";
	String OUT_OFFICE_HOURS="Actual Time spent on Incidence Out of  UK Office (Minutes)";	
	String IN_OFFICE_EFFORT="IN_OFFICE_EFFORT";
	String OUT_OFFICE_EFFORT="OUT_OFFICE_EFFORT";
	void addEeffortsCoumnn(String mainFile) throws  IOException;
	void updateEfforts(String filePath,List<String> files, String startDate, String endDate) throws IOException, ParseException;
	void mergeSheets(List<String> fileNames, String directory, String startDate, String endDate)
			throws IOException;
	
}
