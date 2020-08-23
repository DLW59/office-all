package com.dlw.architecture.office.model.excel;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Map;

/**
 * @author dengliwen
 * @date 2020/6/28
 * @desc 下拉模型
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class DropDownModel {

    /**
     * 单个下拉框数据(非级联)
     */
    private List<String> singleDropDownList;

    /**
     * 级联下拉数据
     */
    private Map<String,List<String>> cascadeDropDownMap;

    /**
     * 依赖的列名
     */
    private String dependFieldName;

}
