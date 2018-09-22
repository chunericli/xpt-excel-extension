package com.xpt.extension.abst;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.xpt.extension.domain.HugeDataExcelOutputDomain;
import com.xpt.extension.exception.ExcelHandleException;

/**
 * 可选择是分批次Excel导出和多Sheet导出;
 * 
 * @author LEric
 */
public abstract class HugeDataExcelExportUtil {

	public void excelHugeDataExport(HugeDataExcelOutputDomain domain) throws Exception {
		boolean multiSheet = domain.isMultiSheets();
		if (multiSheet) {
			handleMulti(domain);
		} else {
			handle(domain);
		}
	}

	public void handleMulti(HugeDataExcelOutputDomain domain) throws IOException {
		int totalRowNums = domain.getTotalRowNums();
		int filterRowNums = domain.getFilterRowNums();
		String fileName = new String(domain.getFileName().getBytes("ISO-8859-1"), Charset.defaultCharset());
		String filePath = domain.getFilePath();

		File file = new File(filePath);
		if (!file.exists()) {
			file.mkdirs();
		}
		if (filterRowNums >= totalRowNums) {
			filterRowNums = totalRowNums;
		}
		int tempsize = (totalRowNums % filterRowNums) == 0 ? totalRowNums / filterRowNums
				: totalRowNums / filterRowNums + 1;

		String tempExcelFile = filePath + fileName + ".xlsx";
		FileOutputStream fos = new FileOutputStream(tempExcelFile);
		// 在内存当中保持 100 行 , 超过的数据放到硬盘中
		SXSSFWorkbook workbook = new SXSSFWorkbook(100);

		for (int i = 0; i < tempsize; i++) {
			Map<String, Object> params = new HashMap<>();
			params.put("limit", i * filterRowNums);
			if (i == (totalRowNums / filterRowNums)) {
				params.put("offset", totalRowNums);
			} else {
				params.put("offset", (i + 1) * filterRowNums);
			}
			List<Map<String, Object>> dataList = getData(params);
			workbook = exportDataToExcel(workbook, dataList, i);
		}

		try {
			workbook.write(fos);
			fos.flush();
		} catch (Exception e) {
			throw new ExcelHandleException(e);
		} finally {
			if (Objects.nonNull(fos)) {
				fos.close();
			}
			if (Objects.nonNull(workbook)) {
				workbook.dispose();
			}
		}
	}

	@SuppressWarnings("resource")
	public void handle(HugeDataExcelOutputDomain domain) throws IOException {
		int totalRowNums = domain.getTotalRowNums();
		int filterRowNums = domain.getFilterRowNums();
		String fileName = new String(domain.getFileName().getBytes("ISO-8859-1"), Charset.defaultCharset());
		String filePath = domain.getFilePath();

		File file = new File(filePath);
		if (!file.exists()) {
			file.mkdirs();
		}
		if (filterRowNums >= totalRowNums) {
			filterRowNums = totalRowNums;
		}
		int tempsize = (totalRowNums % filterRowNums) == 0 ? totalRowNums / filterRowNums
				: totalRowNums / filterRowNums + 1;

		for (int i = 0; i < tempsize; i++) {
			Map<String, Object> params = new HashMap<>();
			params.put("limit", i * filterRowNums);
			if (i == (totalRowNums / filterRowNums)) {
				params.put("offset", totalRowNums);
			} else {
				params.put("offset", (i + 1) * filterRowNums);
			}
			List<Map<String, Object>> dataList = getData(params);
			String tempExcelFile = filePath + fileName + "[" + (i + 1) + "].xlsx";
			FileOutputStream fos = new FileOutputStream(tempExcelFile);

			// 在内存当中保持 100 行 , 超过的数据放到硬盘中
			SXSSFWorkbook workbook = new SXSSFWorkbook(100);
			try {
				workbook = exportDataToExcel(workbook, dataList, i);
				workbook.write(fos);
				fos.flush();
			} catch (Exception e) {
				throw new ExcelHandleException(e);
			} finally {
				if (Objects.nonNull(fos)) {
					fos.close();
				}
				if (Objects.nonNull(workbook)) {
					workbook.dispose();
				}
			}
		}
	}

	public abstract List<Map<String, Object>> getData(Map<String, Object> params);

	public abstract Map<String, Object> getHeaderMetadata();

	/**
	 * 可利用反射动态匹配映射实现
	 * 
	 * @param workbook
	 * @param dataList
	 * @param size
	 * @return
	 */
	public SXSSFWorkbook exportDataToExcel(SXSSFWorkbook workbook, List<Map<String, Object>> dataList, int size) {
		Map<String, Object> objects = getHeaderMetadata();
		String[] headNames = (String[]) objects.get("headNames");
		String[] headEngNames = (String[]) objects.get("headEngNames");
		String sheetName = (String) objects.get("sheetName");

		Cell cell = null;
		Sheet sheet = workbook.createSheet(sheetName);
		Row row = sheet.createRow(0);
		for (int i = 0; i < headNames.length; i++) {
			cell = row.createCell(i);
			cell.setCellValue(headNames[i]);
		}

		if (Objects.nonNull(dataList) && dataList.size() > 0 && !dataList.isEmpty()) {
			int rowIndex = 1;
			for (Map<String, Object> map : dataList) {
				row = sheet.createRow(rowIndex++);
				int index = 0;
				for (int i = 0; i < headEngNames.length; i++) {
					cell = row.createCell(index++);
					cell.setCellType(CellType.STRING);
					cell.setCellValue(("" + map.get(headEngNames[i])).replace("null", ""));
				}
			}
		}
		return workbook;
	}
}