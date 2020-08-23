package com.dlw.architecture.office.annotation;
import com.dlw.architecture.office.enums.ColorType;
import com.dlw.architecture.office.enums.PdfCellType;
import com.dlw.architecture.office.enums.PdfFontType;
import com.itextpdf.text.BaseColor;

import java.lang.annotation.*;

/**
 * @author dengliwen
 * @date 2020/6/17
 * @desc pdf导出字段注解 表示需要导出的列
 * @since 4.0.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface PdfCell {
	/**
	 * 表示列中文名 用于表头
	 * 当设置为有表头的时候name不能为空 必须保持表头与列字段数相等
	 * @return
	 */
	String name() default "";

	/**
	 * 列的索引 表示表头顺序 从0开始
	 * @return
	 */
	int index() default 9999;

	/**
	 * 列的宽度 每列占总宽度的比列
	 * @return
	 */
	float width();

	/**
	 * 单元格类型 默认文本
	 * @return
	 */
	PdfCellType cellType() default PdfCellType.TEXT;

	/**
	 * 单元格样式类型 默认普通文本
	 * @return
	 */
	PdfFontType fontType() default PdfFontType.NORMAL;

	/**
	 * -1表示默认值  如果不使用默认值则设置的值必须大于0
	 * 字体大小 此属性配合fontType一起使用时，如果size=-1表示使用fontType的默认size值，
	 * 否则使用size属性值设置字体大小
	 * @return
	 */
	int size() default -1;

	/**
	 * 设置单元格颜色
	 * @return
	 */
	ColorType cellColor() default ColorType.BLACK;

}
