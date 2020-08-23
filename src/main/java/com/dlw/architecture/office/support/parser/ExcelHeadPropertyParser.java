package com.dlw.architecture.office.support.parser;

import lombok.NoArgsConstructor;

/**
 * @author dengliwen
 * @date 2020/7/6
 * @desc excel head 属性解析
 * @since 4.0.0
 */

@NoArgsConstructor
public class ExcelHeadPropertyParser extends AbstractPropertyParser {

    private static ExcelHeadPropertyParser parser = new ExcelHeadPropertyParser();

    public static ExcelHeadPropertyParser getInstance() {
        return parser;
    }

    public static ExcelHeadPropertyParser getInstance(String prefix, String suffix) {
        return new ExcelHeadPropertyParser(prefix,suffix);
    }

    public ExcelHeadPropertyParser(String prefix, String suffix) {
        super(prefix, suffix);
    }

}
