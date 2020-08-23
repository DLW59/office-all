package com.dlw.architecture.office.annotation;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.*;

/**
 * @author dengliwen
 * @date 2020/6/29
 * @desc 表头样式配置注解
 * @since 4.0.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelHeadStyle {

    /**
     *颜色填充类型
     * @return
     */
    FillPatternType fillPatternType() default FillPatternType.SOLID_FOREGROUND;

    /**
     * 单元格前景填充颜色  默认为黄色
     * @return
     */
    IndexedColors fillBackgroundColor() default IndexedColors.YELLOW;
}
