package com.dlw.architecture.office.support.parser;


import com.dlw.architecture.office.exception.OfficeException;

/**
 * @author dengliwen
 * @date 2020/7/6
 * @desc 属性解析器
 * @since 4.0.0
 */
public interface PropertyParser {

    /**
     * 属性解析
     * @param key 属性key
     * @return 属性key对应的value
     * @throws OfficeException
     */
    String parse(String key) throws OfficeException;
}
