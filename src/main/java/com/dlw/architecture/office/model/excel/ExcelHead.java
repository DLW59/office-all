package com.dlw.architecture.office.model.excel;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * @author dengliwen
 * @date 2020/6/28
 * @desc excel 表头信息
 * @since 4.0.0
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ExcelHead {

    /**
     * 表头对应的列号
     */
    private Integer columnIndex;
    /**
     * 表头对应的字段名称
     */
    private String fieldName;
    /**
     * 表头所有名称
     */
    private List<String> headNameList;

}
