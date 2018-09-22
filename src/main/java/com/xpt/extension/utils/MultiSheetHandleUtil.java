package com.xpt.extension.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

/**
 * 可以利用反射动态映射匹配
 * 
 * @author LEric
 */
public class MultiSheetHandleUtil {

	private MultiSheetHandleUtil() {
	}

	public static Map<String, List<Object>> importData(String filePath, Object object)
			throws FileNotFoundException, IOException, InstantiationException, IllegalAccessException {
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(new File(filePath)));
		HSSFSheet sheet = null;
		Map<String, List<Object>> mapListData = new HashMap<>();
		int sheetNum = workbook.getNumberOfSheets();

		for (int i = 0; i < sheetNum; i++) {
			sheet = workbook.getSheetAt(i);
			List<Object> dataList = new ArrayList<>();
			int rowNum = sheet.getPhysicalNumberOfRows();

			for (int j = 0; j < rowNum; j++) {
				HSSFRow row = sheet.getRow(j);
				int cellNum = row.getPhysicalNumberOfCells();
				Object obj = object.getClass().newInstance();
				Field[] fields = obj.getClass().getDeclaredFields();

				for (int k = 0; k < cellNum; k++) {
					Field field = fields[k];
					field.setAccessible(true);
					field.set(obj, row.getCell(k));
				}
				dataList.add(obj);
			}
			mapListData.put(sheet.getSheetName(), dataList);
		}
		workbook.close();
		return mapListData;
	}

	public static void export(List<String> sheetNameList, Map<String, List<String>> mapSheetHeaderList,
			Map<String, List<Object>> sheetDataList, String filePath) throws Exception {
		HSSFWorkbook workBook = new HSSFWorkbook();
		int sheetNum = sheetNameList.size();
		for (int sn = 0; sn < sheetNum; sn++) {
			HSSFSheet sheet = workBook.createSheet();
			String sheetName = sheetNameList.get(sn);
			workBook.setSheetName(sn, sheetName);

			HSSFRow headerRow = sheet.createRow(0);
			List<String> headerList = mapSheetHeaderList.get(sheetName);
			int headerNum = headerList.size();
			for (int hn = 0; hn < headerNum; hn++) {
				HSSFCell headerCell = headerRow.createCell(hn);
				headerCell.setCellValue(headerList.get(hn));
			}

			int rowIndex = 1;
			List<Object> dataList = sheetDataList.get(sheetName);
			int dataListNum = dataList.size();
			for (int dn = 0; dn < dataListNum; dn++) {
				Object object = dataList.get(dn);
				HSSFRow row = sheet.createRow(rowIndex);
				Field[] fields = object.getClass().getDeclaredFields();
				int fieldLength = fields.length;
				for (int i = 0; i < fieldLength; i++) {
					Field field = fields[i];
					field.setAccessible(true);

					String cellValue = SheetUtil.formmatterCellValue(field, field.get(object));
					Cell cell = row.createCell(i, CellType.STRING);
					cell.setCellValue(cellValue);
				}
				rowIndex++;
			}
		}
		FileOutputStream fos = new FileOutputStream(filePath);
		workBook.write(fos);
		fos.close();
		workBook.close();
	}
}