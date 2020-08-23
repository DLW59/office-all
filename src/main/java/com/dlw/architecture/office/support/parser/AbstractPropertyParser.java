package com.dlw.architecture.office.support.parser;

import com.dlw.architecture.office.exception.OfficeException;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;

/**
 * @author dengliwen
 * @date 2020/7/6
 * @desc 抽象属性解析器
 * @since 4.0.0
 */
@AllArgsConstructor
@NoArgsConstructor
@Getter
public abstract class AbstractPropertyParser implements PropertyParser {

    private String prefix = "${";
    private String suffix = "}";

    @Override
    public String parse(String key) throws OfficeException {
        if (key.startsWith(prefix) && key.endsWith(suffix)) {
            key = key.replace(prefix,"").replace(suffix,"");
            return key;
        }
        throw new OfficeException(String.format("The attribute key【%s】format is incorrect ",key));
    }
}
