package com.dlw.architecture.office.enums;

/**
 * @author dengliwen
 * @date 2020/6/8
 * @desc 文件类型枚举 后续自己扩展
 * @since 4.0.0
 */
public enum FileType {

    EXCEL,
    WORD,
    PDF
    ;

    public enum ExcelType {
        XLS,
        XLSX
    }

    public enum wordType {
        DOC,
        DOCX
    }
}
