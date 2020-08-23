package com.dlw.architecture.office.support.loader;

import com.dlw.architecture.office.support.resource.ResourceLoader;

/**
 * @author dengliwen
 * @date 2020/7/6
 * @desc excel表头属性加载器
 * @since 4.0.0
 */
public class ExcelHeadPropertyLoader extends AbstractPropertyLoader {

    private static ExcelHeadPropertyLoader loader = new ExcelHeadPropertyLoader(new ResourceLoader());

    public static ExcelHeadPropertyLoader getInstance() {
        return loader;
    }

    public ExcelHeadPropertyLoader(ResourceLoader resourceLoader) {
        super(resourceLoader);
    }
}
