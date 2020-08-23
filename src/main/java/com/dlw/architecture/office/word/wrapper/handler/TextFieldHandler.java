package com.dlw.architecture.office.word.wrapper.handler;

import com.dlw.architecture.office.exception.OfficeException;
import org.apache.commons.lang3.StringUtils;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.Map;

/**
 * @author dengliwen
 * @date 2020/6/12
 * @desc 文本字段类型处理器
 * @since 4.0.0
 */
public class TextFieldHandler<T> extends AbstractFieldHandler<T> {

    public TextFieldHandler(Field field) {
        super(field);
    }

    @Override
    public void handle(Map<String, Object> result, Field field, T t, Annotation annotation) throws OfficeException {
        String fieldName = field.getName();
        final Object value = getFieldValue(field, t);
        if (StringUtils.isNotBlank(this.alias)) {
            fieldName = this.alias;
        }
        result.put(fieldName,value);
    }
}
