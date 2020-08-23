package com.dlw.architecture.office.annotation;

import java.lang.annotation.*;

/**
 * @author dengliwen
 * @date 2020/6/12
 * @desc 作用于文本字段
 * @since 4.0.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface WordText {
}
