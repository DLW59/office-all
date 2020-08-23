package com.dlw.architecture.office.word.wrapper;

import com.deepoove.poi.config.Configure;
import com.deepoove.poi.policy.HackLoopTableRenderPolicy;
import com.dlw.architecture.office.annotation.WordLoopTable;
import com.dlw.architecture.office.exception.OfficeException;
import com.dlw.architecture.office.word.util.WordAnnotationUtil;
import com.dlw.architecture.office.word.wrapper.handler.AbstractFieldHandler;
import org.apache.commons.lang3.ArrayUtils;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

/**
 * @author dengliwen
 * @date 2020/6/11
 * @desc word 数据包装类 调用方调wrap接口，把封装的数据对象转换成map结构进行导出
 * 参数必须为BaseWordModel及子类
 * @since 4.0.0
 */
public class WordDataWrapper {

    /**
     * 数据包装处理
     * @param t 传入待处理的数据
     * @param <T> 泛型
     * @return map结构数据 用于直接导出的最终数据
     * @throws OfficeException
     */
    public static <T> Map<String,Object> wrap(T t) throws OfficeException {
        Map<String,Object> result = new HashMap<>();
        final Class<?> clazz = t.getClass();
        final Field[] fields = clazz.getDeclaredFields();
        if (ArrayUtils.isEmpty(fields)) {
            throw new OfficeException(clazz.getName() + " no field defined");
        }
        boolean isLoopTable = false;
        Configure config = Configure.newBuilder().build();
        for (Field field : fields) {
            final Annotation fieldAnnotation = WordAnnotationUtil.getFieldAnnotation(field);
            if (fieldAnnotation == null) {
                continue;
            }
            field.setAccessible(true);
            AbstractFieldHandler.wrapField(result,fieldAnnotation,field,t);
            if (field.isAnnotationPresent(WordLoopTable.class)) {
                isLoopTable = true;
                config.customPolicy(field.getName(),new HackLoopTableRenderPolicy());
            }
        }
        if (isLoopTable) {
            result.put(WordLoopTable.class.getName(),config);
        }
        return result;
    }
}
