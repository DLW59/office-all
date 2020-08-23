package com.dlw.architecture.office.pdf;

import com.dlw.architecture.office.annotation.PdfCell;
import com.dlw.architecture.office.annotation.PdfTable;
import com.dlw.architecture.office.enums.ColorType;
import com.dlw.architecture.office.enums.PdfCellType;
import com.dlw.architecture.office.enums.PdfFontType;
import com.dlw.architecture.office.exception.OfficeException;
import com.dlw.architecture.office.pdf.style.PdfStyle;
import com.dlw.architecture.office.util.ClassUtil;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.Field;
import java.net.URL;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @author dengliwen
 * @date 2020/6/16
 * @desc pdf创建工具类
 * @since 4.0.0
 */
@Slf4j
public class PdfUtil {

    /**
     * 创建PDF表格
     * @param clazz 表数据实体类
     * @param data 导出的数据
     * @return table
     */
    public static PdfPTable createTable(Class<?> clazz, Collection<?> data) throws Exception {
        if (null == clazz) {
            throw new OfficeException("No pdf class specified");
        }
        List<Field> fieldList = getFields(clazz);
        final float[] widths = getWidths(fieldList);
        List<String> headers = getHeaders(fieldList);
        final PdfTable pdfTable = clazz.getAnnotation(PdfTable.class);
        if (null == pdfTable) {
            return null;
        }
        tableValidate(fieldList, headers, pdfTable);
        return createTable(widths,headers,pdfTable,data);
    }

    /**
     * PDF table 注解属性校验
     * @param fieldList 列字段
     * @param headers 表头
     * @param pdfTable 表
     */
    private static void tableValidate(List<Field> fieldList, List<String> headers, PdfTable pdfTable) throws OfficeException {
        if (pdfTable.isHead()) {
            if (fieldList.size() != headers.size()) {
                throw new OfficeException("the number of headers and the number of columns are not equal");
            }
            if (pdfTable.headSize() != -1 && pdfTable.headSize() <= 0) {
                throw new OfficeException("The table head size does not use the default value and is less than 0");
            }
        }
        if (pdfTable.isLockWidth() && pdfTable.totalWidth() <= 0) {
            throw new OfficeException("in lock width mode, the total width of the table cannot be less than 0");
        }
        final int size = pdfTable.titleSize();
        if (size != -1 && size <= 0) {
            throw new OfficeException("The table title size does not use the default value and is less than 0");
        }
    }

    /**
     * 创建表格
     * @param widths 列宽
     * @param headers 表头
     * @param pdfTable 表格
     * @param data 数据
     * @return
     * @throws Exception
     */
    private static PdfPTable createTable(float[] widths, List<String> headers, PdfTable pdfTable, Collection<?> data)
            throws Exception {
        final String title = pdfTable.title();
        PdfPTable table = new PdfPTable(widths);
        if (pdfTable.isLockWidth()) {
            table.setTotalWidth(pdfTable.totalWidth());
            table.setLockedWidth(true);
        }
        table.getDefaultCell().setHorizontalAlignment(Element.ALIGN_CENTER);
        table.getDefaultCell().setVerticalAlignment(Element.ALIGN_CENTER);
        if (StringUtils.isNotBlank(title)) {
            //创建标题 不加边框
            table.addCell(mergeCell(pdfTable,title, widths.length));
        }
        final boolean isHead = pdfTable.isHead();
        if (isHead) {
            createHeadCell(headers, table,pdfTable);
        }
        fillTableCellData(data,table,isHead);
        return table;
    }

    /**
     * 创建表头单元格
     * @param headers 表头
     * @param table 表格
     * @param pdfTable 表注解属性
     */
    private static void createHeadCell(List<String> headers, PdfPTable table, PdfTable pdfTable) throws OfficeException {
        for (String header : headers) {
            table.addCell(createHeadCellWithBorder(header,pdfTable));
        }
    }

    /**
     * table 填充数据
     * @param data 数据
     * @param table 表格
     * @param isHead 是否有表头
     * @throws Exception
     */
    private static void fillTableCellData(Collection<?> data, PdfPTable table, boolean isHead) throws Exception {
        for (Object o : data) {
            final List<Field> fields = getFields(o.getClass());
            for (Field field : fields) {
                addTableCell(table, isHead, o, field);
            }
        }
    }

    /**
     * 添加表格单元格数据
     * @param table 表格
     * @param isHead 是否带表头
     * @param o 行数据
     * @param field 字段
     * @throws Exception
     */
    private static void addTableCell(PdfPTable table, boolean isHead, Object o, Field field) throws Exception {
        field.setAccessible(true);
        Object value = getFieldValue(o, field);
        final PdfCell pdfCell = field.getAnnotation(PdfCell.class);
        final PdfCellType cellType = pdfCell.cellType();
        final PdfFontType fontType = pdfCell.fontType();
        int size = pdfCell.size();
        switch (cellType) {
            case TEXT:
                createCellWithText(table, isHead, (String) value,fontType,size,pdfCell.cellColor());
                break;
            case IMAGE:
                createCellWithImage(table, value);
                break;
            default:
                throw new OfficeException("unsupported cell type,see " + PdfCellType.class.getName());
        }
    }

    /**
     * 创建文本类单元格
     * @param table 表格
     * @param isHead 是否有表头
     * @param value 单元格值
     * @param fontType 文本格式
     * @param size
     * @param colorType
     */
    private static void createCellWithText(PdfPTable table, boolean isHead, String value, PdfFontType fontType, int size,
                                           ColorType colorType) throws OfficeException {
        PdfPCell cell;
        cell = createCell(isHead,value,fontType,size, colorType);
        table.addCell(cell);
    }

    /**
     * 创建图片类单元格
     * @param table 表
     * @param value 单元格值
     * @throws Exception
     */
    private static void createCellWithImage(PdfPTable table, Object value) throws Exception {
        Image image;
        if (value instanceof String) {
            image = Image.getInstance((String) value);
        }else if (value instanceof URL) {
            image = Image.getInstance((URL) value);
        }else if (value instanceof byte[]) {
            image = Image.getInstance((byte[]) value);
        }else {
            throw new OfficeException("unsupported pdf image field type");
        }
        table.addCell(image);
    }
    /**
     * 获取字段值 并转化一些类型数据
     * @param o 字段对象
     * @param field 字段
     * @return 字段值
     * @throws IllegalAccessException
     */
    private static Object getFieldValue(Object o, Field field) throws IllegalAccessException {
        Class<?> type = field.getType();
        Object value = field.get(o);
        if (value == null) {
            value = "";
        } else if (type == Date.class || type == Timestamp.class) {
            Date date = type == Date.class ? (Date)value : new Date(((Timestamp)value).getTime());
            value = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(date);
        } else {
            value = String.valueOf(value);
        }
        return value;
    }

    /**
     * 合并单元格
     *
     * @param pdfTable
     * @param title 标题
     * @param colspan 合并的列数
     * @return
     */
    private static PdfPCell mergeCell(PdfTable pdfTable, String title, int colspan) throws OfficeException {
        PdfPCell cell = createCell(false,title,pdfTable.titleFontType(),pdfTable.titleSize(),pdfTable.titleColor());
        cell.setHorizontalAlignment(Element.ALIGN_LEFT);
        //将该单元格所在行包括该单元格在内的colspan列单元格合并为一个单元格
        cell.setColspan(colspan);
        cell.setPadding(3.0f);
        cell.setPaddingTop(15.0f);
        cell.setPaddingBottom(8.0f);
        return cell;
    }

    /**
     * 创建表格标题 带边框
     * @param head 表头
     * @param pdfTable 头注解
     * @return
     */
    private static PdfPCell createHeadCellWithBorder(String head,PdfTable pdfTable) throws OfficeException {
        final PdfFontType fontType = pdfTable.headFontType();
        int size = pdfTable.headSize();
        return createCell(true,head,fontType,size, pdfTable.headColor());
    }

    /**
     * 创建表格单元格 带边框
     *
     * @param isHead 是否有表头
     * @param content 单元格内容
     * @param fontType 样式类型
     * @param size 字体大小
     * @param colorType 颜色类型
     * @return
     */
    private static PdfPCell createCell(boolean isHead, String content, PdfFontType fontType, int size, ColorType colorType) throws OfficeException {
        PdfPCell cell = new PdfPCell();
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        if (size == -1) {
            size = fontType.getSize();
        }
        Phrase phrase;
        switch (fontType) {
            case NORMAL:
                phrase = new Phrase(content, PdfStyle.getTextFont(size,colorType));
                break;
            case BOLD:
                phrase = new Phrase(content, PdfStyle.getBoldFont(size,colorType));
                break;
            case UNDERLINE:
                phrase = new Phrase(content, PdfStyle.getUnderlineFont(size,colorType));
                break;
            case FIRST_TITLE:
                phrase = new Phrase(content, PdfStyle.getFirstTitleFont(size,colorType));
                break;
            case SECOND_TITLE:
                phrase = new Phrase(content, PdfStyle.getSecondTitleFont(size,colorType));
                break;
            default:
                throw new OfficeException("unsupported font style type");
        }
        cell.setPhrase(phrase);
        if (!isHead) {
            cell.setBorder(0);
        }
        return cell;
    }

    /**
     * 获取所有指定注解的字段且根据注解属性排序
     * @param clazz
     * @return
     */
    private static List<Field> getFields(Class<?> clazz) throws OfficeException {
        final Field[] fields = ClassUtil.getAnnotationFields(clazz, PdfCell.class);
        List<Field> fieldList = Arrays.asList(fields);
        for (Field field : fieldList) {
            if (field.getAnnotation(PdfCell.class).index() < 0) {
                throw new OfficeException("@PdfColumn annotated index attribute value cannot be less than 0");
            }
        }
        fieldList.sort((o1, o2) -> {
            final int index1 = o1.getAnnotation(PdfCell.class).index();
            final int index2 = o2.getAnnotation(PdfCell.class).index();
            return index1 - index2;
        });
        return fieldList;
    }

    /**
     * 获取PDF 所有表头名称
     * @param fieldList 字段
     * @return
     */
    private static List<String> getHeaders(List<Field> fieldList) {
        return fieldList.stream().map(field -> field.getAnnotation(PdfCell.class).name())
                .filter(StringUtils::isNotBlank)
                .collect(Collectors.toList());
    }

    /**
     * 获取PDF 列的宽度
     * @param fieldList 列字段
     * @return
     */
    private static float[] getWidths(List<Field> fieldList) throws OfficeException {
        float[] widths = new float[fieldList.size()];
        for (int i = 0; i < fieldList.size(); i++) {
            final PdfCell pdfCell = fieldList.get(i).getAnnotation(PdfCell.class);
            if (pdfCell.width() <= 0) {
                throw new OfficeException("The width of the table cell is less than 0");
            }
            widths[i] = pdfCell.width();
            if (pdfCell.size() != -1 && pdfCell.size() <= 0) {
                throw new OfficeException("The table cell size does not use the default value and is less than 0");
            }
        }
        return widths;
    }
}
