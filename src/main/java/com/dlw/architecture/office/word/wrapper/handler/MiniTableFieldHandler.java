package com.dlw.architecture.office.word.wrapper.handler;

import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.dlw.architecture.office.annotation.WordMiniTable;
import com.dlw.architecture.office.exception.OfficeException;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.*;

/**
 * @author dengliwen
 * @date 2020/6/12
 * @desc 带表头表格字段类型处理器
 * @since 4.0.0
 */
public class MiniTableFieldHandler<T> extends AbstractFieldHandler<T> {

    public MiniTableFieldHandler(Field field) {
        super(field);
    }

    @Override
    public void handle(Map<String, Object> result, Field field, T t, Annotation annotation) throws OfficeException {
        WordMiniTable miniTable = (WordMiniTable) annotation;
        String fieldName = field.getName();
        if (StringUtils.isNotBlank(this.alias)) {
            fieldName = this.alias;
        }
        final Object value = getFieldValue(field, t);
        if (Objects.isNull(value)) {
            return;
        }

        final String[] header = miniTable.header();
        if (ArrayUtils.isEmpty(header)) {
            throw new OfficeException("The table header cannot be empty");
        }
        List<RowRenderData> body = new ArrayList<>();
        convertFieldValue(value, header, body);
        List<TextRenderData> headerData = new ArrayList<>(header.length);
        for (String head : header) {
            headerData.add(new TextRenderData(head));
        }
        final RowRenderData buildHead = RowRenderData.build(headerData.toArray(new TextRenderData[0]));
        final MiniTableRenderData miniTableRenderData = new MiniTableRenderData(buildHead,body);
        result.put(fieldName,miniTableRenderData);
    }

    private void convertFieldValue(Object value, String[] header, List<RowRenderData> body) throws OfficeException {
        if (value instanceof Collection) {
            Collection collection = (Collection) value;
            if (collection.isEmpty()) {
                return;
            }
            final Object next = collection.iterator().next();
            final Field[] fields = next.getClass().getDeclaredFields();
            if (ArrayUtils.isEmpty(fields)) {
                throw new OfficeException("Table entity has no defined fields");
            }
            if (header.length != fields.length) {
                throw new OfficeException("The number of header and column fields are not equal");
            }
            for (Object o : collection) {
                final Field[] declaredFields = o.getClass().getDeclaredFields();
                List<String> fieldValues = new ArrayList<>(declaredFields.length);
                for (Field declaredField : declaredFields) {
                    try {
                        declaredField.setAccessible(true);
                        final Object fieldValue = declaredField.get(o);
                        if (isPrimitive(fieldValue) || fieldValue instanceof String) {
                            fieldValues.add(String.valueOf(fieldValue));
                        }else {
                            throw new OfficeException("@WordMiniTable Annotated object field must be of Primitive type or String");
                        }
                    } catch (IllegalAccessException e) {
                        throw new OfficeException("Failed to get field value",e);
                    }
                }
                final String[] array = fieldValues.toArray(new String[0]);
                body.add(RowRenderData.build(array));
            }
        }else {
            throw new OfficeException("@WordMiniTable annotation currently only supports Collection type fields");
        }
    }
}
