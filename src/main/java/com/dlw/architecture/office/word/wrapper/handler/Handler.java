package com.dlw.architecture.office.word.wrapper.handler;


import com.dlw.architecture.office.exception.OfficeException;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.Map;

/**
 * @author dengliwen
 * @date 2020/6/12
 * @desc 字段处理器
 * @since 4.0.0
 */
@FunctionalInterface
public interface Handler<T> {

    /**
     * 字段处理方法
     * @param result 字段值处理后的结果
     * @param field 处理的字段
     * @param t 操作的数据对象(调用端封装传过来的数据模型)
     * @param annotation 字段的word注解
     * @throws OfficeException
     */
    void handle(Map<String, Object> result, Field field, T t, Annotation annotation) throws OfficeException;
}
