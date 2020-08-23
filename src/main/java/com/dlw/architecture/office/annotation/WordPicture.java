package com.dlw.architecture.office.annotation;

import java.lang.annotation.*;

/**
 * @author dengliwen
 * @date 2020/6/12
 * @desc 作用于图片字段
 * @since 4.0.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface WordPicture {

    /**
     * 图片宽度
     * @return
     */
    int width() default 200;

    /**
     * 图片高度
     * @return
     */
    int height() default 200;

    /**
     * 图片格式
     * @return
     */
    String format() default ".png";
}
