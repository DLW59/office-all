package com.dlw.architecture.office.util;

import com.dlw.architecture.office.exception.OfficeException;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.beanutils.BeanUtils;

import java.beans.BeanInfo;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author dengliwen
 * @date 2020/6/16
 * @desc 对象与map的转换
 * @since 4.0.0
 */
@Slf4j
public class ConvertUtil {

    /**
     * 对象转为map
     * @param bean 对象
     * @return map
     * @throws OfficeException
     */
    public static Map<String,Object> beanToMap(Object bean) throws OfficeException {
        Map<String, Object> map = new HashMap<>();
        try {
            //获取JavaBean的描述器
            BeanInfo b = Introspector.getBeanInfo(bean.getClass(), Object.class);
            //获取属性描述器
            PropertyDescriptor[] pds = b.getPropertyDescriptors();
            //对属性迭代
            for (PropertyDescriptor pd : pds) {
                //属性名称
                String propertyName = pd.getName();
                //属性值,用getter方法获取
                Method m = pd.getReadMethod();
                //用对象执行getter方法获得属性值
                Object properValue = m.invoke(bean);

                //把属性名-属性值 存到Map中
                map.put(propertyName, properValue);
            }
        }catch (Exception e) {
            throw new OfficeException("bean convert to map failed " + e.getMessage(),e);
        }
        return map;
    }

    /**
     * nao转换bean
     * @param rowMap map对象
     * @param clazz 目标class
     * @param <T> 泛型
     * @return
     */
    private static <T> T mapToBean(Map<String, Object> rowMap, Class<T> clazz) throws OfficeException {
        if (rowMap == null) {
            return null;
        }
        T obj;
        try {
            obj = clazz.newInstance();
            BeanUtils.populate(obj, rowMap);
        } catch (Exception e) {
            throw new OfficeException("map convert to object failed",e);
        }
        return obj;
    }


    /**
     * 转换为对象
     * @param headers 表头
     * @param fields 表头字段
     * @param sheetData sheet数据
     * @param clazz 转换的class
     * @param <T>
     * @return
     */
    public static <T> List<T> convertExcelDataToBean(String[] headers, Field[] fields,
                                                     List<String[]> sheetData,Class<T> clazz) throws OfficeException {
        List<T> excelData = new ArrayList<>();
        for (String[] row : sheetData) {
            Map<String,Object> rowMap = new HashMap<>();
            for (int i = 0; i < headers.length; i++) {
                rowMap.put(fields[i].getName(),row[i]);
            }
            excelData.add(mapToBean(rowMap,clazz));
        }
        return excelData;
    }

}
