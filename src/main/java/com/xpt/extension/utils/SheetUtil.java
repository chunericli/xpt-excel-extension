package com.xpt.extension.utils;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Objects;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

import com.xpt.extension.annotation.ExcelCell;
import com.xpt.extension.annotation.ExcelSheet;
import com.xpt.extension.domain.ExcelImportDomain;
import com.xpt.extension.exception.ExcelHandleException;

public class SheetUtil {

	private SheetUtil() {
	}

	/**
	 * 数据量最好限定在10万左右，要不然可能存在性能问题
	 * 
	 * @param sheet
	 * @param domain
	 * @return
	 * @throws IllegalAccessException
	 * @throws InstantiationException
	 */
	public static List<Object> getResults(Sheet sheet, ExcelImportDomain domain)
			throws InstantiationException, IllegalAccessException {
		if (Objects.isNull(sheet)) {
			throw new ExcelHandleException("Make sure that excel sheet exists!");
		}

		List<Field> listFields = new ArrayList<>();
		Map<String, Field> mapFields = new HashMap<>();
		Class<?> classDomain = domain.getClass();
		Field[] fields = classDomain.getDeclaredFields();

		// 通常经常下只考虑public，private，protected修饰的field；类似static，final，volatile等修饰的不考虑
		if (Objects.nonNull(fields) && fields.length > 0) {
			for (Field field : fields) {
				listFields.add(field);
				ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
				if (Objects.isNull(excelCell) || Objects.isNull(excelCell.name())
						|| excelCell.name().trim().length() == 0) {
					throw new ExcelHandleException("Make sure that ExcelCell anotation&&name exists!");
				}
				mapFields.put(excelCell.name(), field);
			}
		} else {
			throw new ExcelHandleException("Make sure that domain fields exists!");
		}

		// 解析Excel时校验Domain对象和Excel列名是否匹配，以Excel的列名为参照解析
		ExcelSheet excelSheet = classDomain.getAnnotation(ExcelSheet.class);
		boolean check = excelSheet.check();
		int rowCount = excelSheet.count();

		// Excel列名称的获取，存在单元格的合并只获取Excel最后行的标题
		int rowIndex = 0;
		Map<String, String> matchFields = new HashMap<>();
		for (Row row : sheet) {
			if (rowIndex <= rowCount) {
				for (Cell cell : row) {
					cell.setCellType(CellType.STRING);
					matchFields.put(CellReference.convertNumToColString(cell.getColumnIndex()),
							cell.getStringCellValue());
				}
				rowIndex++;
				continue;
			}
			break;
		}

		// 校验Domain对象和Excel列名是否匹配
		if (check) {
			if (Objects.nonNull(matchFields) && matchFields.size() == mapFields.size()) {
				for (Entry<String, String> entry : matchFields.entrySet()) {
					if (mapFields.containsKey(entry.getValue())) {
						throw new ExcelHandleException("Make sure that domain fields matched excel-columns!");
					}
				}
			} else {
				throw new ExcelHandleException("Make sure that domain fields matched excel-columns!");
			}
		}

		// 映射对象数据返回
		Iterator<Row> sheetIterator = sheet.rowIterator();
		List<Object> dataList = new ArrayList<>();
		while (sheetIterator.hasNext()) {
			Row row = sheetIterator.next();
			if (rowIndex > rowCount && Objects.nonNull(row)) {
				Object domainObj = classDomain.newInstance();
				int size = listFields.size();

				for (int i = 0; i < size; i++) {
					// cell为空表示excel格式“异常”，最常见的是数据删除时保留excel格式
					Cell cell = row.getCell(i);
					if (Objects.isNull(cell)) {
						// 【“空数据”待处理】
						continue;
					}

					String mapValue = matchFields.get(CellReference.convertNumToColString(cell.getColumnIndex()));
					Field mapField = mapFields.get(mapValue);
					if (Objects.isNull(mapField)) {
						continue;
					}
					Object cellValue = getCellValue(mapField, cell);
					mapField.setAccessible(true);
					mapField.set(domainObj, cellValue);
				}
				dataList.add(domainObj);
			}
		}
		return dataList;
	}

	@SuppressWarnings("deprecation")
	private static Object getCellValue(Field field, Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			return cell.getRichStringCellValue().getString();
		case Cell.CELL_TYPE_NUMERIC:
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				// return cell.getDateCellValue());
				ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
				if (Objects.isNull(excelCell) || Objects.isNull(excelCell.dateFormatter())
						|| excelCell.dateFormatter().trim().length() == 0) {
					throw new ExcelHandleException("Make sure that ExcelCell anotation&&dateFormatter exists!");
				}
				Date date = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
				SimpleDateFormat sdf = new SimpleDateFormat(excelCell.dateFormatter().trim());
				return sdf.format(date);
			} else {
				return cell.getNumericCellValue();
			}
		case Cell.CELL_TYPE_BOOLEAN:
			return cell.getBooleanCellValue();
		case Cell.CELL_TYPE_FORMULA:
			return cell.getCellFormula();
		case Cell.CELL_TYPE_BLANK:
			return null;
		case Cell.CELL_TYPE_ERROR:
			return null;
		default:
			return null;
		}
	}

	public static String formmatterCellValue(Field field, Object fieldValue) {
		if (Objects.isNull(fieldValue)) {
			return null;
		}
		Class<?> fieldType = field.getType();

		if (Date.class.equals(fieldType)) {
			ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
			if (Objects.isNull(excelCell) || Objects.isNull(excelCell.dateFormatter())
					|| excelCell.dateFormatter().trim().length() == 0) {
				throw new ExcelHandleException("Make sure that ExcelCell anotation&&dateFormatter exists!");
			}
			SimpleDateFormat dateFormat = new SimpleDateFormat(excelCell.dateFormatter().trim());
			return dateFormat.format(fieldValue);
		} else {
			return String.valueOf(fieldValue);
		}
	}
}