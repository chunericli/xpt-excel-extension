package com.xpt.extension.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ ElementType.FIELD })
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelCell {

	/**
	 * 单元格的名称
	 *
	 * @return
	 */
	String name() default "";

	/**
	 * 日期格式的解析
	 * 
	 * @return
	 */
	String dateFormatter() default "yyyy-MM-dd hh:mm:ss";
}