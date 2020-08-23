package com.dlw.architecture.office.model.excel;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @author dengliwen
 * @date 2020/6/28
 * @desc 单元格合并数据模型
 * @since 4.0.0
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class CellRange {

    /**
     * 起始行号
     */
    private int firstRow;
    /**
     * 结束行号
     */
    private int lastRow;
    /**
     * 起始列号
     */
    private int firstCol;
    /**
     * 结束列号
     */
    private int lastCol;
}
