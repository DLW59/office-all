package com.dlw.architecture.office.word.wrapper.handler;

import com.deepoove.poi.data.NumbericRenderData;
import com.deepoove.poi.data.RenderData;
import com.deepoove.poi.data.TextRenderData;
import com.dlw.architecture.office.annotation.WordList;
import com.dlw.architecture.office.exception.OfficeException;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.*;

/**
 * @author dengliwen
 * @date 2020/6/12
 * @desc 列表字段类型处理器
 * @since 4.0.0
 */
public class ListFieldHandler<T> extends AbstractFieldHandler<T> {

    private static Map<Integer, Pair<STNumberFormat.Enum, String>> formatMap = new HashMap<>();

    /**
     * 初始化列表格式
     */
    static {
        formatMap.put(0, NumbericRenderData.FMT_BULLET);
        formatMap.put(1, NumbericRenderData.FMT_DECIMAL);
        formatMap.put(2, NumbericRenderData.FMT_DECIMAL_PARENTHESES);
        formatMap.put(3, NumbericRenderData.FMT_LOWER_LETTER);
        formatMap.put(4, NumbericRenderData.FMT_LOWER_ROMAN);
        formatMap.put(5, NumbericRenderData.FMT_UPPER_LETTER);
        formatMap.put(6, NumbericRenderData.FMT_UPPER_ROMAN);
    }

    public ListFieldHandler(Field field) {
        super(field);
    }

    @Override
    public void handle(Map<String, Object> result, Field field, T t, Annotation annotation) throws OfficeException {
        WordList wordList = (WordList) annotation;
        String fieldName = field.getName();
        if (StringUtils.isNotBlank(this.alias)) {
            fieldName = this.alias;
        }
        final Object value = getFieldValue(field, t);
        if (Objects.isNull(value)) {
            return;
        }
        final List<? extends RenderData> renderData = convertFieldValue(value);
        final int format = wordList.format();
        final Pair<STNumberFormat.Enum, String> fmt = formatMap.getOrDefault(format, NumbericRenderData.FMT_BULLET);
        final NumbericRenderData numbericRenderData = new NumbericRenderData(fmt,renderData);
        result.put(fieldName,numbericRenderData);
    }

    /**
     * 转换字段值为导出的数据类型
     * @param value 字段的值
     * @return
     */
    private List<? extends RenderData> convertFieldValue(Object value) throws OfficeException {
        List<RenderData> renderData = new ArrayList<>();
        if (value instanceof Collection) {
            Collection collection = (Collection) value;
            if (collection.isEmpty()) {
                return renderData;
            }
            for (Object o : collection) {
                //取集合里的元素判断类型
                if (isPrimitive(o) || o instanceof String) {
                    RenderData data = new TextRenderData(String.valueOf(o));
                    renderData.add(data);
                }else {
                    throw new OfficeException("@WordList Annotated object field must be of Primitive type or String");
                }
            }
        }else if (value instanceof String) {
            renderData.add(new TextRenderData((String) value));
        }else {
            throw new OfficeException("@WordList annotation currently only supports String and Collection type fields");
        }
        return renderData;
    }

}
