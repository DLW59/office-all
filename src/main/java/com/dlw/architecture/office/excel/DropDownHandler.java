package com.dlw.architecture.office.excel;


import com.dlw.architecture.office.model.excel.DropDownModel;
import com.dlw.architecture.office.model.excel.ExcelHead;

/**
 * @author dengliwen
 * @date 2020/6/28
 * @desc 下拉数据
 * @since 4.0.0
 *
 */
@FunctionalInterface
public interface DropDownHandler {

    /**
     * 封装下拉数据
     * @param excelHead excel 表头信息
     * @return 下拉框数据
     */
    DropDownModel dropDown(ExcelHead excelHead);
}
