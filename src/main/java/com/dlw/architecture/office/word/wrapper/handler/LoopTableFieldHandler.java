package com.dlw.architecture.office.word.wrapper.handler;

import com.deepoove.poi.config.Configure;
import com.deepoove.poi.policy.HackLoopTableRenderPolicy;
import com.dlw.architecture.office.exception.OfficeException;
import org.apache.commons.lang3.StringUtils;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.Map;
import java.util.Objects;

/**
 * @author dengliwen
 * @date 2020/6/12
 * @desc 循环表格字段类型处理器
 * @since 4.0.0
 */
public class LoopTableFieldHandler<T> extends AbstractFieldHandler<T> {

    public LoopTableFieldHandler(Field field) {
        super(field);
    }

    @Override
    public void handle(Map<String, Object> result, Field field, T t, Annotation annotation) throws OfficeException {
        String fieldName = field.getName();
        if (StringUtils.isNotBlank(this.alias)) {
            fieldName = this.alias;
        }
        final Object value = getFieldValue(field, t);
        if (Objects.isNull(value)) {
            return;
        }
        result.put(fieldName,value);
    }
}
