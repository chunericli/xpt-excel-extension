package com.xpt.extension.utils;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;
import java.util.Objects;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.xpt.extension.exception.ExcelHandleException;

import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.transformer.XLSTransformer;

/**
 * <jx:forEach items="${?}" var="?"> </jx:forEach>；<jx:if test="${?}" > </jx:if>
 * 
 * $[sum(position)]；$[min(position)]；$[max(position)]
 * 
 * $[average(position)]；$[count(position)]
 * 
 * @author LEric
 */
public class TemplateExcelExportUtil {

	private TemplateExcelExportUtil() {
	}

	public static void export(String templatePath, Map<String, Object> params, String fileName,
			HttpServletResponse response) throws Exception {
		XLSTransformer transformer = new XLSTransformer();
		InputStream is = TemplateExcelExportUtil.class.getResourceAsStream(templatePath);
		OutputStream out = response.getOutputStream();
		try {
			response.setHeader("Content-Disposition",
					"attachment;filename=\"" + new String(fileName.getBytes("ISO-8859-1"), "UTF-8"));
			response.setContentType("application/vnd.ms-excel");
			transformer.transformXLS(is, params).write(out);
			out.flush();
		} catch (Exception e) {
			throw new ExcelHandleException(e.getCause());
		} finally {
			if (Objects.nonNull(out)) {
				out.close();
			}
			if (Objects.nonNull(is)) {
				is.close();
			}
		}
	}

	/**
	 * 导出数据到文件系统
	 * 
	 * @param source
	 * @param params
	 * @param target
	 * @throws Exception
	 */
	public static void export(String source, Map<String, Object> params, String target) {
		XLSTransformer transformer = new XLSTransformer();
		try {
			transformer.transformXLS(source, params, target);
		} catch (ParsePropertyException | InvalidFormatException | IOException e) {
			throw new ExcelHandleException(e.getCause());
		}
	}
}