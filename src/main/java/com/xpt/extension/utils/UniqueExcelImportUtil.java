package com.xpt.extension.utils;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Objects;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.xpt.extension.domain.ExcelImportDomain;
import com.xpt.extension.exception.ExcelHandleException;
import com.xpt.extension.mapping.ExcelSheet;

/**
 * 只支持单sheet页的导入
 * 
 * @author LEric
 */
public class UniqueExcelImportUtil {

	private UniqueExcelImportUtil() {
	}

	/**
	 * 只考虑简单的filePath，例如：/opt/pmo/object.wps
	 *
	 * @param filePath
	 * @param domain
	 * @return
	 * @throws IOException
	 * @throws InvalidFormatException
	 * @throws EncryptedDocumentException
	 * @throws IllegalAccessException
	 * @throws InstantiationException
	 */
	public static List<Object> importExcelByPath(String filePath, ExcelImportDomain domain)
			throws EncryptedDocumentException, InvalidFormatException, IOException, InstantiationException,
			IllegalAccessException {
		return importExcelByFile(FileUtil.checkGetFile(filePath), domain);
	}

	public static List<Object> importExcelByFile(File excelFile, ExcelImportDomain domain)
			throws EncryptedDocumentException, InvalidFormatException, IOException, InstantiationException,
			IllegalAccessException {
		Workbook workbook = WorkbookFactory.create(excelFile);
		return importExcelByWorkbook(workbook, domain);
	}

	public static List<Object> importExcel(InputStream inputStream, ExcelImportDomain domain)
			throws EncryptedDocumentException, InvalidFormatException, IOException, InstantiationException,
			IllegalAccessException {
		Workbook workbook = WorkbookFactory.create(inputStream);
		return importExcelByWorkbook(workbook, domain);
	}

	@SuppressWarnings("unchecked")
	public static List<Object> importExcelByWorkbook(Workbook workbook, ExcelImportDomain domain)
			throws InstantiationException, IllegalAccessException {
		return excelImportHandle(workbook, (Class<ExcelImportDomain>) domain.getClass());
	}

	private static List<Object> excelImportHandle(Workbook workbook, Class<ExcelImportDomain> domain)
			throws InstantiationException, IllegalAccessException {
		int sheetNum = workbook.getNumberOfSheets();
		if (sheetNum <= 0) {
			throw new ExcelHandleException("Make sure that file format correct!");
		}

		ExcelSheet excelSheet = domain.getAnnotation(ExcelSheet.class);
		if (Objects.isNull(excelSheet)) {
			throw new ExcelHandleException("Make sure that ExcelSheet annotation exists!");
		}

		Sheet sheet = null;
		if (Objects.nonNull(excelSheet.name()) && excelSheet.name().trim().length() > 0) {
			sheet = workbook.getSheet(excelSheet.name().trim());
		} else {
			sheet = workbook.getSheetAt(0);
		}
		return SheetUtil.getResults(sheet, domain.newInstance());
	}
}