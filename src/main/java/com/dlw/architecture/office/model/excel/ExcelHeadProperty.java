package com.dlw.architecture.office.model.excel;

import com.dlw.architecture.office.annotation.ExcelCell;
import com.dlw.architecture.office.annotation.ExcelHeadStyle;
import com.dlw.architecture.office.annotation.ExcelSheet;
import com.dlw.architecture.office.excel.ExcelHelper;
import com.dlw.architecture.office.excel.style.ExcelStyleHelper;
import com.dlw.architecture.office.exception.OfficeException;
import com.dlw.architecture.office.support.loader.AbstractPropertyLoader;
import com.dlw.architecture.office.support.loader.ExcelHeadPropertyLoader;
import com.dlw.architecture.office.support.parser.AbstractPropertyParser;
import com.dlw.architecture.office.support.parser.ExcelHeadPropertyParser;
import com.dlw.architecture.office.util.ClassUtil;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.TreeMap;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @author dengliwen
 * @date 2020/6/28
 * @desc excel 表头属性,包含获取表头的对象class，行号，表头，表头样式
 * @since 4.0.0
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ExcelHeadProperty {

    /**
     * class 用于获取表头、字段等excel信息
     */
    private Class<?> headClazz;
    /**
     * 行号
     */
    private int headRowNumber;
    /**
     * 表头信息
     */
    private Map<Integer, ExcelHead> headMap = new TreeMap<>();

    /**
     * 表头样式
     */
    private Map<Integer, CellStyle> headStyleMap = new TreeMap<>();

    private AbstractPropertyLoader loader = ExcelHeadPropertyLoader.getInstance();

    private AbstractPropertyParser parser;

    public ExcelHeadProperty(Class<?> headClazz) {
        this.headClazz = headClazz;
    }

    /**
     * 初始化表头信息
     * @param wb
     */
    public void initHead(Workbook wb) throws OfficeException {
        int i = 0;
        loadProperties();
        initParser();
        for (Field field : ExcelHelper.getAnnotationSortFields(headClazz)) {
            String[] names = field.getAnnotation(ExcelCell.class).name();
            final String[] finalNames = parseName(names);
            headMap.put(i, new ExcelHead(i, field.getName(), Stream.of(finalNames).collect(Collectors.toList())));
            initHeadStyle(wb, i, field);
            i++;
        }
        initHeadRowNumber();
    }

    /**
     * 加载excel head配置文件
     */
    private void loadProperties() throws OfficeException {
        try {
            final ExcelSheet excelSheet = headClazz.getAnnotation(ExcelSheet.class);
            if (null == excelSheet) {
                return;
            }
            final String[] configPath = excelSheet.headConfigPath();
            if (ArrayUtils.isEmpty(configPath)) {
                return;
            }
            loader.load(configPath);
        } catch (Exception e) {
            throw new OfficeException("Failed to load configuration file",e);
        }
    }

    /**
     * 初始化解析器
     */
    private void initParser() {
        final ExcelSheet excelSheet = ClassUtil.getAnnotation(headClazz, ExcelSheet.class);
        if (null != excelSheet) {
            final String prefix = excelSheet.prefix();
            final String suffix = excelSheet.suffix();
            if (StringUtils.isNotBlank(prefix) && StringUtils.isNotBlank(suffix)) {
                parser = ExcelHeadPropertyParser.getInstance(prefix,suffix);
                return;
            }
        }
        parser = ExcelHeadPropertyParser.getInstance();
    }
    /**
     * 解析excel head的属性
     * @param names 表头名称
     * @return 表头
     */
    private String[] parseName(String[] names) throws OfficeException {
        if (names.length != 1) {
            return names;
        }
        final String prefix = parser.getPrefix();
        final String suffix = parser.getSuffix();
        if (names[0].startsWith(prefix) && names[0].endsWith(suffix)) {
            final String key = parser.parse(names[0]);
            final Properties properties = loader.getProperties();
            final String property = properties.getProperty(key);
            if (StringUtils.isBlank(property)) {
                throw new OfficeException(String.format("Attribute key【%s】 not found",key));
            }
            return property.split(",");
        }
        return names;
    }




    /**
     * 初始化表头样式
     * @param wb 当前工作簿
     * @param i 表头列号
     * @param field 表头字段
     */
    private void initHeadStyle(Workbook wb, int i, Field field) {
        final ExcelHeadStyle headStyle = field.getAnnotation(ExcelHeadStyle.class);
        if (null == headStyle) {
            headStyleMap.put(i, ExcelStyleHelper.createDefaultHeadStyle(wb));
        }else {
            headStyleMap.put(i, ExcelStyleHelper.createHeadStyle(wb,headStyle));
        }
    }


    /**
     * 初始化行号
     */
    private void initHeadRowNumber() {
        for (ExcelHead head : headMap.values()) {
            List<String> list = head.getHeadNameList();
            if (list != null && list.size() > headRowNumber) {
                headRowNumber = list.size();
            }
        }
        for (ExcelHead head : headMap.values()) {
            List<String> list = head.getHeadNameList();
            if (list != null && !list.isEmpty() && list.size() < headRowNumber) {
                int lack = headRowNumber - list.size();
                int last = list.size() - 1;
                for (int i = 0; i < lack; i++) {
                    list.add(list.get(last));
                }
            }
        }
    }
}
