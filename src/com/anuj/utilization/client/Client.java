package com.anuj.utilization.client;

import java.io.IOException;

import com.anuj.utilization.client.impl.UtilizationReportImpl;
import com.anuj.utilization.client.inf.UtilizationReport;

public class Client {
	public static void main(String[] args) throws IOException {
		
		

		String utilizationFolder = "C:\\Users\\pankaja\\Desktop\\L2\\automation\\timesheet\\Sept\\week5\\";
		//dd-MM-yyyy
		String startDate ="21-09-2019";
		String endDate ="27-09-2019";
		
	/*	String utilizationFolder =null;
		String startDate =null;
		String endDate = null;
		
		if(args!=null && args.length<3) {
			throw new IOException("Missing arguments to run this util");
		}
		utilizationFolder = args[0];
		 startDate = args[1];
		 endDate = args[2];*/
		UtilizationReport utilizationReport = new UtilizationReportImpl(utilizationFolder,startDate,endDate);

		utilizationReport.prepareUtilizatoinReport();

	}
	
} 