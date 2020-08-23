package com.dlw.architecture.office.annotation;

import java.lang.annotation.*;

/**
 * @author dengliwen
 * @date 2020/6/12
 * @desc 作用于列表字段
 * @since 4.0.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface WordList {

    /** 列表格式 对应关系如下 为方便简化为数字代替
     * 0 -> FMT_BULLET //● ● ●
     * 1 -> FMT_DECIMAL //1. 2. 3.
     * 2 -> FMT_DECIMAL_PARENTHESES //1) 2) 3)
     * 3 -> FMT_LOWER_LETTER //a. b. c.
     * 4-> FMT_LOWER_ROMAN //i ⅱ ⅲ
     * 5 -> FMT_UPPER_LETTER //A. B. C.
     * 6 -> FMT_UPPER_ROMAN //I. II. III.
     * @return
     */
    int format() default 0;
}
