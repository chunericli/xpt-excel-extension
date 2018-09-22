package com.xpt.extension.utils;

import java.io.BufferedOutputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.xpt.extension.domain.ExcelOutputDomain;
import com.xpt.extension.exception.ExcelHandleException;
import com.xpt.extension.mapping.ExcelCell;
import com.xpt.extension.mapping.ExcelSheet;

/**
 * 只支持单sheet页的导出
 * 
 * @author LEric
 */
public class UniqueExcelExportUtil {

	private UniqueExcelExportUtil() {
	}

	/**
	 * 导出数据到文件系统
	 * 
	 * @param filePath
	 * @param domain
	 * @throws IOException
	 */
	public static void exportExcelToFileByPath(String filePath, List<ExcelOutputDomain> domain) throws IOException {
		Workbook workbook = obtainWorkbookByDomain(domain);
		exportExcelToFileByWorkbook(filePath, workbook);
	}

	public static void exportExcelToFileByWorkbook(String filePath, Workbook workbook) throws IOException {
		FileOutputStream fileOutputStream = null;
		try {
			fileOutputStream = new FileOutputStream(filePath);
			workbook.write(fileOutputStream);
			fileOutputStream.flush();
		} catch (Exception e) {
			throw new ExcelHandleException(e.getCause());
		} finally {
			try {
				if (Objects.nonNull(fileOutputStream)) {
					fileOutputStream.close();
				}
				if (Objects.nonNull(workbook)) {
					workbook.close();
				}
			} catch (Exception e) {
				new ExcelHandleException(e.getCause());
			}
		}
	}

	/**
	 * HSSFworkbook：操作Excel2003版本，扩展名为xls;
	 * 
	 * XSSFworkbook：操作Excel2007版本，扩展名为xlsx;
	 * 
	 * SXSSFworkbook ：用于大数据量导出，allowable range 65536(0..65535)
	 * 
	 * @param domain
	 * @return
	 * @throws IOException
	 */
	public static Workbook obtainWorkbookByDomain(List<ExcelOutputDomain> domain) throws IOException {
		if (Objects.isNull(domain) || domain.isEmpty() || domain.size() == 0) {
			throw new ExcelHandleException("Make sure that domain data exists!");
		}
		Workbook workbook = new HSSFWorkbook();
		obtainSheetByDomain(workbook, domain);
		return workbook;
	}

	public static void obtainSheetByDomain(Workbook workbook, List<ExcelOutputDomain> domain) {
		if (Objects.isNull(domain) || domain.isEmpty() || domain.size() == 0) {
			throw new ExcelHandleException("Make sure that domain data exists!");
		}

		Class<? extends ExcelOutputDomain> classDomain = domain.get(0).getClass();
		ExcelSheet excelSheet = classDomain.getAnnotation(ExcelSheet.class);
		if (Objects.isNull(excelSheet)) {
			throw new ExcelHandleException("Make sure that ExcelSheet annotation exists!");
		}
		String sheetName = null;
		if (Objects.nonNull(excelSheet.name()) && excelSheet.name().trim().length() > 0) {
			sheetName = excelSheet.name().trim();
		} else {
			sheetName = classDomain.getSimpleName();
		}

		Sheet sheet = workbook.createSheet(sheetName);
		Field[] fields = classDomain.getDeclaredFields();
		List<Field> listFields = new ArrayList<>();

		// 通常经常下只考虑public，private，protected修饰的field；类似static，final，volatile等修饰的不考虑
		if (Objects.nonNull(fields) && fields.length > 0) {
			for (Field field : fields) {
				listFields.add(field);
			}
		} else {
			throw new ExcelHandleException("Make sure that domain fields exists!");
		}

		int fieldSize = listFields.size();
		Row headRow = sheet.createRow(0);
		for (int i = 0; i < fieldSize; i++) {
			Field field = listFields.get(i);

			ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
			if (Objects.isNull(excelCell) || Objects.isNull(excelCell.name())
					|| excelCell.name().trim().length() == 0) {
				throw new ExcelHandleException("Make sure that ExcelCell anotation&&name exists!");
			}
			String fieldName = excelCell.name().trim();

			Cell cell = headRow.createCell(i, CellType.STRING);
			cell.setCellValue(String.valueOf(fieldName));
		}

		int dataSize = domain.size();
		for (int ds = 0; ds < dataSize; ds++) {
			Object data = domain.get(ds);
			Row row = sheet.createRow(ds + 1);

			for (int fs = 0; fs < fieldSize; fs++) {
				Field field = listFields.get(fs);
				try {
					field.setAccessible(true);
					String cellValue = SheetUtil.formmatterCellValue(field, field.get(data));
					Cell cell = row.createCell(fs, CellType.STRING);
					cell.setCellValue(cellValue);
				} catch (IllegalAccessException e) {
					throw new RuntimeException(e);
				}
			}
		}
	}

	public static byte[] exportToBytes(List<ExcelOutputDomain> domain) throws IOException {
		Workbook workbook = obtainWorkbookByDomain(domain);
		ByteArrayOutputStream byteArrayOutputStream = null;
		try {
			byteArrayOutputStream = new ByteArrayOutputStream();
			workbook.write(byteArrayOutputStream);
			byteArrayOutputStream.flush();
			return byteArrayOutputStream.toByteArray();
		} catch (Exception e) {
			throw new ExcelHandleException(e.getCause());
		} finally {
			try {
				if (Objects.nonNull(byteArrayOutputStream)) {
					byteArrayOutputStream.close();
				}
				if (Objects.nonNull(workbook)) {
					workbook.close();
				}
			} catch (Exception e) {
				throw new ExcelHandleException(e.getCause());
			}
		}
	}

	public static void exportExcelToBrowserByDomain(HttpServletResponse response, List<ExcelOutputDomain> domain,
			String fileName) throws IOException, IllegalAccessException, ClassNotFoundException {
		Workbook workbook = obtainWorkbookByDomain(domain);
		exportExcelToBrowserByWorkbook(response, workbook, fileName);
	}

	public static void exportExcelToBrowserByWorkbook(HttpServletResponse response, Workbook workbook, String fileName)
			throws IOException, IllegalAccessException, ClassNotFoundException {
		OutputStream outputStream = null;
		BufferedOutputStream bufferedOutputStream = null;
		try {
			// 文件名称有后缀
			response.reset();
			response.setHeader("Content-Disposition", "attachment;filename=" + fileName);
			response.setContentType("application/vnd.ms-excel;charset=UTF-8");
			response.setHeader("Pragma", "no-cache");
			response.setHeader("Cache-Control", "no-cache");
			response.setDateHeader("Expires", 0);

			outputStream = response.getOutputStream();
			bufferedOutputStream = new BufferedOutputStream(outputStream);
			bufferedOutputStream.flush();
			workbook.write(bufferedOutputStream);
		} catch (Exception e) {
			throw new ExcelHandleException(e.getCause());
		} finally {
			if (Objects.nonNull(bufferedOutputStream)) {
				bufferedOutputStream.close();
			}
			if (Objects.nonNull(outputStream)) {
				outputStream.close();
			}
			if (Objects.nonNull(workbook)) {
				workbook.close();
			}
		}
	}

	public static void exportExcelToBrowserByFilePath(HttpServletResponse response, String filePath)
			throws IOException, IllegalAccessException, ClassNotFoundException {
		OutputStream outputStream = null;
		BufferedOutputStream bufferedOutputStream = null;
		InputStream inputStream = null;
		try {
			// 文件名称有后缀
			response.reset();
			response.setHeader("Content-Disposition",
					"attachment;filename=" + filePath.substring(filePath.lastIndexOf(File.separator) + 1));
			response.setContentType("application/vnd.ms-excel;charset=UTF-8");
			response.setHeader("Pragma", "no-cache");
			response.setHeader("Cache-Control", "no-cache");
			response.setDateHeader("Expires", 0);

			inputStream = new FileInputStream(new File(filePath));
			outputStream = response.getOutputStream();
			bufferedOutputStream = new BufferedOutputStream(outputStream);

			bufferedOutputStream.flush();

			int length = 0;
			byte[] buffer = new byte[1024];
			while ((length = inputStream.read(buffer)) != -1) {
				bufferedOutputStream.write(buffer, 0, length);
			}
		} catch (Exception e) {
			throw new ExcelHandleException(e.getCause());
		} finally {
			if (Objects.nonNull(bufferedOutputStream)) {
				bufferedOutputStream.close();
			}
			if (Objects.nonNull(outputStream)) {
				outputStream.close();
			}
			if (Objects.nonNull(inputStream)) {
				inputStream.close();
			}
		}
	}
}