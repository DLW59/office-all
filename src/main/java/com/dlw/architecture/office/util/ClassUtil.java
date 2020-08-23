package com.dlw.architecture.office.util;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @author dengliwen
 * @date 2020/6/18
 * @desc
 * @since 4.0.0
 */
public class ClassUtil {

    /**
     * 获取class的字段 包括父类的
     * @param clazz
     * @return
     */
    public static Field[] getClassFields(Class<?> clazz) {
        List<Field> list = new ArrayList<>();
        Field[] fields;
        do {
            fields = clazz.getDeclaredFields();
            Collections.addAll(list, fields);
            clazz = clazz.getSuperclass();
        } while (clazz != Object.class && clazz != null);
        return list.toArray(fields);
    }

    /**
     * 获取有指定注解的字段
     * @param aClass 目标class
     * @param clazz 注解class
     * @return
     */
    public static Field[] getAnnotationFields(Class aClass,Class<? extends Annotation> clazz) {
        List<Field> fieldList = arrayToList(getClassFields(aClass));
        fieldList = fieldList.stream().filter(field -> Objects.nonNull(field.getAnnotation(clazz)))
                .collect(Collectors.toList());
        return fieldList.toArray(new Field[0]);
    }


    public static List<Field> arrayToList(Field[] fields) {
        return Stream.of(fields).collect(Collectors.toList());
    }

    /**
     * 获取指定class的指定注解
     * @param clazz  对象class
     * @param aClass 注解class
     * @param <T> 注解类型
     * @return 注解
     */
    public static <T extends Annotation> T getAnnotation(Class<?> clazz, Class<T> aClass) {
        final T t = clazz.getAnnotation(aClass);
        return t;
    }
}
