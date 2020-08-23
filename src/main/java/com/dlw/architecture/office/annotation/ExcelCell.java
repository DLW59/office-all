package com.dlw.architecture.office.annotation;
import com.alibaba.excel.annotation.ExcelProperty;

import java.lang.annotation.*;

/**
 * @author dengliwen
 * @date 2020/6/17
 * @desc excel导出字段注解 表示需要导出的列
 * @since 4.0.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelCell {
	/**
	 * 表示列中文名 作用于表头
	 * name保持唯一
	 * 支持多级表头
	 * @return
	 */
	String[] name();

	/**
	 * 列的索引 表示表头顺序 从0开始
	 * @return
	 */
	int index() default 9999;

	/**
	 * excel cell宽度
	 * @return
	 */
	int width() default 20;
}
