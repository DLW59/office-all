package com.dlw.architecture.office.word.util;

import com.dlw.architecture.office.annotation.*;
import com.dlw.architecture.office.exception.OfficeException;
import org.apache.commons.lang3.ArrayUtils;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.List;

/**
 * @author dengliwen
 * @date 2020/6/14
 * @desc word 注解工具 用于获取字段上的word类型的注解
 * @since 4.0.0
 */
public class WordAnnotationUtil {
    private static List<Class<? extends Annotation>> wordFieldAnnotation;

    static {
        wordFieldAnnotation = Arrays.asList(
                WordList.class,
                WordLoopTable.class,
                WordMiniTable.class,
                WordPicture.class,
                WordText.class);
    }

    /**
     * word字段注解校验 一个字段只能有一个word类型注解
     *
     * @param field
     */
    public static Annotation getFieldAnnotation(Field field) throws OfficeException {
        final Annotation[] annotations = field.getDeclaredAnnotations();
        if (ArrayUtils.isEmpty(annotations)) {
            return null;
        }
        Annotation an = null;
        if (annotations.length > 1) {
            int count = 0;
            for (Annotation annotation : annotations) {
                if (wordFieldAnnotation.contains(annotation.annotationType())) {
                    an = annotation;
                    count++;
                }
            }
            if (count > 1) {
                throw new OfficeException("annotation of more than one word type field");
            }
        }else {
            if (wordFieldAnnotation.contains(annotations[0].annotationType())) {
                an = annotations[0];
            }
        }
        return an;
    }
}
