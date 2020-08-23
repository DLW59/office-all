package com.dlw.architecture.office.word.wrapper.handler;

import com.deepoove.poi.el.Name;
import com.dlw.architecture.office.annotation.*;
import com.dlw.architecture.office.exception.OfficeException;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.Map;
import java.util.Objects;

/**
 * @author dengliwen
 * @date 2020/6/12
 * @desc 字段处理器抽象类。各种word字段处理的基类 提供一些公共方法
 * <p>AbstractFieldHandler(Field field)</p> 该构造方法用于初始化字段别名
 * @since 4.0.0
 */
public abstract class AbstractFieldHandler<T> implements Handler<T> {

    /**
     * 字段别名 对应模板的标签
     */
    protected String alias;

    public AbstractFieldHandler(Field field) {
        final Name name = field.getDeclaredAnnotation(Name.class);
        if (Objects.nonNull(name)) {
            //字段别名 真正对应模板的标签
            this.alias = name.value();
        }
    }

    /**
     * 获取字段值
     * @param field 字段
     * @return 字段值
     */
    protected Object getFieldValue(Field field,T t) throws OfficeException {
        final Object value;
        try {
            value = field.get(t);
        } catch (IllegalAccessException e) {
            throw new OfficeException("access field value is failed",e);
        }
        return value;
    }

    /**
     * 根据字段注解类型 选择字段的处理器
     * @param result 处理后的数据结构
     * @param fieldAnnotation 字段的注解
     * @param field 字段
     * @param t 传入的数据对象
     * @param <T> 泛型参数
     * @throws OfficeException 没有找到word相应的字段注解会抛该异常
     */
    public static <T> void wrapField(Map<String,Object> result, Annotation fieldAnnotation, Field field, T t) throws OfficeException {
        AbstractFieldHandler abstractFieldHandler;
        if (fieldAnnotation.annotationType().equals(WordText.class)) {
            abstractFieldHandler = new TextFieldHandler(field);
        }else if (fieldAnnotation.annotationType().equals(WordPicture.class)) {
            abstractFieldHandler = new PictureFieldHandler(field);
        }else if (fieldAnnotation.annotationType().equals(WordList.class)) {
            abstractFieldHandler = new ListFieldHandler(field);
        }else if (fieldAnnotation.annotationType().equals(WordMiniTable.class)) {
            abstractFieldHandler = new MiniTableFieldHandler(field);
        }else if (fieldAnnotation.annotationType().equals(WordLoopTable.class)) {
            abstractFieldHandler = new LoopTableFieldHandler(field);
        }else {
            throw new OfficeException("Unsupported word field annotation");
        }
        abstractFieldHandler.handle(result,field,t,fieldAnnotation);
    }

    /**
     * 是否为基本类型
     * @param o
     * @return
     */
    protected boolean isPrimitive(Object o) {
        boolean flag;
        try {
            flag = ((Class) o.getClass().getField("TYPE").get(null)).isPrimitive();
        } catch (IllegalAccessException | NoSuchFieldException e) {
            flag = false;
        }
        return flag;
    }
}
