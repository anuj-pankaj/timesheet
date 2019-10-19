package com.anuj.utilization.client.inf;

public interface MergingFilesService {
	
	void mergeIncidentSheet(String directory);
	void mergeNonIncidentSheet(String filePath);
	void mergeOthertask(String filePath);
	void mergeCTask(String filePath);

}
