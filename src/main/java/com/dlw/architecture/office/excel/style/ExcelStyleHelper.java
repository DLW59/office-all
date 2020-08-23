package com.dlw.architecture.office.excel.style;

import com.dlw.architecture.office.annotation.ExcelHeadStyle;
import org.apache.poi.ss.usermodel.*;

/**
 * @author dengliwen
 * @date 2020/6/29
 * @desc excel 样式辅助类 提供表头和内容的默认样式方法，支持自定义表头属性设置
 * @since 4.0.0
 */
public class ExcelStyleHelper {

    /**
     * 创建表头默认的单元格格式
     *
     * @param wb 工作簿
     * @return CellStyle
     */
    public static CellStyle createDefaultHeadStyle(Workbook wb) {

        CellStyle cs = createDefaultStyle(wb);
        // 单元格前景色填充
        cs.setFillForegroundColor(IndexedColors.YELLOW.index);
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return cs;
    }

    /**
     * 根据表头注解设置样式
     * @param wb  工作簿
     * @param headStyle 表头样式注解
     * @return 样式
     */
    public static CellStyle createHeadStyle(Workbook wb, ExcelHeadStyle headStyle) {
        CellStyle cs = createDefaultStyle(wb);
        cs.setFillForegroundColor(headStyle.fillBackgroundColor().index);
        cs.setFillPattern(headStyle.fillPatternType());
        return cs;
    }
    /**
     * 创建内容默认的单元格格式
     *
     * @param wb 工作簿
     * @return CellStyle
     */
    public static CellStyle createDefaultContentStyle(Workbook wb) {
        CellStyle cs = createDefaultStyle(wb);
        return cs;
    }

    /**
     * 创建默认的单元格格式
     *
     * @param wb 工作簿
     * @return CellStyle
     */
    private static CellStyle createDefaultStyle(Workbook wb) {
        CellStyle cs = wb.createCellStyle();
        // 字体
        Font f = wb.createFont();
        f.setFontHeightInPoints((short) 14);
        f.setColor(IndexedColors.BLACK.getIndex());
        f.setFontName("黑体");
        cs.setFont(f);
        DataFormat format = wb.createDataFormat();
        cs.setDataFormat(format.getFormat("@"));
        // 边框
        cs.setBorderLeft(BorderStyle.THIN);
        cs.setBorderRight(BorderStyle.THIN);
        cs.setBorderTop(BorderStyle.THIN);
        cs.setBorderBottom(BorderStyle.THIN);
        // 位置
        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);

        //自动换行
        cs.setWrapText(true);
        return cs;
    }

}
