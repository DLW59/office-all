package com.dlw.architecture.office.annotation;

import com.deepoove.poi.data.DocxRenderData;

import java.lang.annotation.*;

/**
 * @author dengliwen
 * @date 2020/6/12
 * @desc 作用于表字段 生成表格 适用于没有表格结构，需要指定表头
 * 对应框架的MiniTableRenderData  @WordMiniTable 注解修饰的字段如果是
 * 对象，则此对象的字段要定义为String类型
 *
 * @since 4.0.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface WordMiniTable {

    /**
     * 表头 根据表头生成表格
     * @return
     */
    String[] header();
}
