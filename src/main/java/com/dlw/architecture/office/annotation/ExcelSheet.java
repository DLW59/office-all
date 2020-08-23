package com.dlw.architecture.office.annotation;

import java.lang.annotation.*;

/**
 * @author dengliwen
 * @date 2020/6/28
 * @desc excel sheet 信息
 * @since 4.0.0
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelSheet {

    /**
     * sheet名称
     * @return
     */
    String name();

    /**
     * excel表头配置文件路径
     * @return
     */
    String[] headConfigPath() default {};

    /**
     * 前缀  默认为 "${"
     * @return
     */
    String prefix() default "";

    /**
     * 后缀 默认为 "}"
     * @return
     */
    String suffix() default "";
}
