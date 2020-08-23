package com.dlw.architecture.office.excel;

import com.dlw.architecture.office.exception.OfficeException;
import com.dlw.architecture.office.model.excel.ExcelHeadProperty;
import com.dlw.architecture.office.util.ConvertUtil;
import com.dlw.architecture.office.util.IoUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.*;
import java.lang.reflect.Field;
import java.util.*;

/**
 * @author dengliwen
 * @date 2020/6/17
 * @desc excel上传工具
 * @since 4.0.0
 */
@Slf4j
public class ExcelImporter {

    /**
     * 读取excel文件数据
     * @param inputStream 文件流
     * @param clazz class信息
     * @param <T>
     * @return 数据
     * @throws OfficeException
     */
    public static <T> List<T> read(InputStream inputStream,Class<T> clazz) throws OfficeException {
        final String sheetName = ExcelHelper.validateClass(clazz);
        return doRead(inputStream, clazz, sheetName);
    }



    /**
     * 根据文件流读取excel文件数据并转换成指定类型集合
     * @param inputStream excel文件流
     * @param clazz 转换的目标类class
     * @param sheetName sheet名称
     * @param <T> 泛型参数类型
     * @return excel数据
     */
    public static <T> List<T> read(InputStream inputStream,Class<T> clazz,String sheetName) throws OfficeException {
        if (StringUtils.isBlank(sheetName)) {
            sheetName = ExcelHelper.validateClass(clazz);
        }
        return doRead(inputStream, clazz, sheetName);
    }


    /**
     * 根据文件流读取excel文件数据并转换成指定类型集合
     * @param inputStream excel文件流
     * @param className 转换的目标类名
     * @param sheetName sheet名称
     * @return excel数据
     */
    public static List<?> read(InputStream inputStream,String className,String sheetName) throws OfficeException {
        Class<?> aClass;
        try {
            aClass = Class.forName(className);
        } catch (ClassNotFoundException e) {
            throw new OfficeException(className + " the specified class was not found " ,e);
        }
        return read(inputStream,aClass,sheetName);
    }

    /**
     * 根据文件流读取excel文件数据并转换成指定类型集合
     * @param inputStream excel文件流
     * @param className 转换的目标类名
     * @return excel数据
     */
    public static <T> List<T> read(InputStream inputStream,String className) throws OfficeException {
        Class<T> aClass;
        try {
            aClass = (Class<T>) Class.forName(className);
        } catch (ClassNotFoundException e) {
            throw new OfficeException(className + " the specified class was not found " ,e);
        }
        return read(inputStream,aClass);
    }


    /**
     * 读取excel的所有sheet数据
     * @param inputStream 文件流数据
     * @param classes 待转换成的对象class
     * @return excel 数据
     */
    public static List<List<?>> read(InputStream inputStream,List<Class<?>> classes) throws OfficeException {
        if (CollectionUtils.isEmpty(classes)) {
           throw new OfficeException("No class specified for conversion object");
        }
        try {
            ByteArrayOutputStream outputStream = IoUtil.reuseInputStream(inputStream);
            List<List<?>> result = new ArrayList<>(classes.size());
            for (int i = 0; i < classes.size(); i++) {
                final Class<?> aClass = classes.get(i);
                final String sheetName = ExcelHelper.validateClass(aClass);
                final List<?> list = doRead(new ByteArrayInputStream(outputStream.toByteArray()), aClass, sheetName);
                result.add(list);
            }
            return result;
        }catch (Exception e) {
            throw new OfficeException("import excel failed:" + e.getMessage(), e);
        }
    }

    /**
     * 真正读取excel的操作 读取指定的sheet数据 目前只适用于单行表头
     * @param inputStream 文件流
     * @param clazz 转换的目标类class
     * @param sheetName sheet名称
     * @param <T> 泛型参数
     * @return sheet内容
     */
    private static <T> List<T> doRead(InputStream inputStream, Class<T> clazz, String sheetName) throws OfficeException {
        Workbook wb;
        try {
            wb = WorkbookFactory.create(inputStream);
            Sheet sheet = ExcelHelper.checkAndGetSheet(wb, sheetName);
            final ExcelHeadProperty headProperty = ExcelHelper.getHeaders(wb, clazz);
            final int headRowNumber = headProperty.getHeadRowNumber();
            final List<String[]> headers = ExcelHelper.readSheetRowData(sheet, 0, headRowNumber - 1);
            final Field[] fields = ExcelHelper.getAnnotationSortFields(clazz);
            final String[] lastHeader = headers.get(headRowNumber - 1);
            final List<String[]> sheetData = ExcelHelper.readSheetData(sheet, headRowNumber - 1);
            List<T> excelData = ConvertUtil.convertExcelDataToBean(lastHeader,fields,sheetData,clazz);
            return excelData;
        } catch (Exception e) {
            throw new OfficeException("import excel failed:" + e.getMessage(),e);
        }
    }
}
