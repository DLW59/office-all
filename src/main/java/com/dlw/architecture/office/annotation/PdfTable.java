package com.dlw.architecture.office.annotation;


import com.dlw.architecture.office.enums.ColorType;
import com.dlw.architecture.office.enums.PdfFontType;

import java.lang.annotation.*;

/**
 * @author dengliwen
 * @date 2020/6/18
 * @desc 设置PDF实体的属性
 * @since 4.0.0
 */
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface PdfTable {

    /**
     * 主题 比如每个pdfTable可以设置不同的主题
     * @return
     */
    String title() default "";

    /**
     * 设置标题颜色
     * @return
     */
    ColorType titleColor() default ColorType.BLACK;

    /**
     * 标题字体类型 默认加粗
     * @return
     */
    PdfFontType titleFontType() default PdfFontType.BOLD;

    /**
     * -1表示默认值  如果不使用默认值则设置的值必须大于0
     * 字体大小 此属性配合titleFontType一起使用时，如果titleSize=-1表示使用titleFontType的默认size值，
     * 否则使用titleSize属性值设置字体大小
     * @return
     */
    int titleSize() default -1;

    /**
     * 表头样式类型
     * @return
     */
    PdfFontType headFontType() default PdfFontType.NORMAL;

    /**
     * -1表示默认值  如果不使用默认值则设置的值必须大于0
     * 字体大小 此属性配合headFontType一起使用时，如果headSize=-1表示使用headFontType的默认size值，
     * 否则使用headSize属性值设置字体大小
     * @return
     */
    int headSize() default -1;

    /**
     * 设置表头颜色
     * @return
     */
    ColorType headColor() default ColorType.BLACK;

    /**
     * 总宽度  表示表格总宽度
     * 如果不是固定宽度模式则可以不设置
     * @return
     */
    float totalWidth() default 0;

    /**
     * 是否固定宽度 默认不固定 自适应宽度
     * 固定宽度需自己控制宽度 如果为true则totalWidth属性值必须大于0
     * @return
     */
    boolean isLockWidth() default false;

    /**
     * 是否有表头 默认有
     * @return
     */
    boolean isHead() default true;
}
