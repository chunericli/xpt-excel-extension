package com.xpt.extension.exception;

@SuppressWarnings("serial")
public class ExcelHandleException extends RuntimeException {

	public ExcelHandleException(String errMsg) {
		super(errMsg);
	}

	public ExcelHandleException(String errMsg, Throwable cause) {
		super(errMsg, cause);
	}
}