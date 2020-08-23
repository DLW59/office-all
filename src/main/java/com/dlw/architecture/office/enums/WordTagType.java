package com.dlw.architecture.office.enums;

/**
 * @author dengliwen
 * @date 2020/6/11
 * @desc word 内容类型
 * @since 4.0.0
 */
public enum WordTagType {
    /**
     * 文本 {{var}}
     */
    TEXT,

    /**
     * 图片 {{@var}}
     */
    PICTURE,

    /**
     * 表格 {{#var}}
     */
    TABLE,

    /**
     * 列表 {{*var}}
     */
    LIST,

    /**
     * 文档 {{+var}}
     */
    DOC,
    /**
     * 超链接 属于文本的一种{{var}}
     */
    LINK

}
