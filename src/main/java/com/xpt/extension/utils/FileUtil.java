package com.xpt.extension.utils;

import java.io.File;

import com.xpt.extension.exception.ExcelHandleException;

public class FileUtil {

	private FileUtil() {
	}

	public static File getFile(String filePath) {
		return new File(filePath);
	}

	public static boolean checkFile(File file) {
		return file.exists();
	}

	public static File checkGetFile(String filePath) {
		File file = getFile(filePath);
		if (!checkFile(file)) {
			throw new ExcelHandleException("Make sure that file exists!");
		}
		return file;
	}
}