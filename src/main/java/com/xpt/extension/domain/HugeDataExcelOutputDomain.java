package com.xpt.extension.domain;

import java.io.Serializable;

@SuppressWarnings("serial")
public class HugeDataExcelOutputDomain implements Serializable {

	private int totalRowNums;

	private int filterRowNums = 50000;

	private String fileName;

	private String filePath;

	private boolean multiSheets = false;

	public boolean isMultiSheets() {
		return multiSheets;
	}

	public void setMultiSheets(boolean multiSheets) {
		this.multiSheets = multiSheets;
	}

	public String getFilePath() {
		return filePath;
	}

	public void setFilePath(String filePath) {
		this.filePath = filePath;
	}

	public String getFileName() {
		return fileName;
	}

	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	public int getTotalRowNums() {
		return totalRowNums;
	}

	public void setTotalRowNums(int totalRowNums) {
		this.totalRowNums = totalRowNums;
	}

	public int getFilterRowNums() {
		return filterRowNums;
	}

	public void setFilterRowNums(int filterRowNums) {
		this.filterRowNums = filterRowNums;
	}
}