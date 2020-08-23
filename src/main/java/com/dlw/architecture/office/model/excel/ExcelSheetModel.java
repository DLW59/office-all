package com.dlw.architecture.office.model.excel;

import com.dlw.architecture.office.excel.DropDownHandler;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Map;

/**
 * @author dengliwen
 * @date 2020/6/17
 * @desc excel 导出数据模型 支持多sheet导出
 * 一个ExcelModel对象对应一个sheet的数据
 * @since 4.0.0
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ExcelSheetModel<T> {

    /**
     * sheet名称
     */
    private String sheetName;

    /**
     * sheet索引 多个sheet用于排序
     */
    private int index;

    /**
     * excel导出模型对应的class 用于获取表头等属性
     */
    private Class<?> clazz;

    /**
     * 该sheet的数据
     */
    private List<T> data;

    /**
     * 下拉接口处理
     */
    private DropDownHandler dropDownHandler;

}
