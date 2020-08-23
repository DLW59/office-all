package com.dlw.architecture.office.annotation;

import java.lang.annotation.*;

/**
 * @author dengliwen
 * @date 2020/6/12
 * @desc 作用于表字段 循环表内容适用于模板有表结构只需填充内容
 * 对应框架的HackLoopTableRenderPolicy
 * @since 4.0.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface WordLoopTable {

}
