package com.dlw.architecture.office.enums;

import lombok.Getter;

/**
 * @author dengliwen
 * @date 2020/6/19
 * @desc pdf样式枚举
 * @since 4.0.0
 */
public enum PdfFontType {

    /**
     * 默认文本
     */
    NORMAL(12),
    /**
     * 加粗
     */
    BOLD(12),
    /**
     * 下划线
     */
    UNDERLINE(12),
    /**
     * 一级标题
     */
    FIRST_TITLE(22),
    /**
     * 二级标题
     */
    SECOND_TITLE(15);

    @Getter
    private int size;

    PdfFontType(int size) {
        this.size = size;
    }
}
