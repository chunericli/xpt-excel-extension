package com.xpt.extension.mapping;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ ElementType.TYPE })
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelSheet {

	/**
	 * 当前Excel Sheet的名称
	 *
	 * @return
	 */
	String name() default "";

	/**
	 * Excel开始解析行数
	 *
	 * @return
	 */
	int count() default 0;

	/**
	 * true表示单元格的名称与Domain对象不匹配时抛出异常
	 * 
	 * @return
	 */
	boolean check() default false;

	/**
	 * 日期格式的解析
	 * 
	 * @return
	 */
	String dateFormatter() default "yyyy-MM-dd hh:mm:ss";
}