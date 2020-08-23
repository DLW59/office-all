package com.kun.sdk.office.test.example;


import com.dlw.architecture.office.excel.DropDownHandler;
import com.dlw.architecture.office.model.excel.DropDownModel;
import com.dlw.architecture.office.model.excel.ExcelHead;

import java.util.*;

/**
 * @author dengliwen
 * @date 2020/6/29
 * @desc
 */
public class MyDropDownHandler implements DropDownHandler {

    @Override
    public DropDownModel dropDown(ExcelHead excelHead) {
        DropDownModel drop = new DropDownModel();
        if (excelHead.getFieldName().equals("city")) {
            drop.setDependFieldName("province");
            Map<String, List<String>> cascadeDropDownMap = new HashMap<>();
            cascadeDropDownMap.put("四川",Arrays.asList("成都","自贡"));
            cascadeDropDownMap.put("云南",Arrays.asList("昆明","大理"));
            drop.setCascadeDropDownMap(cascadeDropDownMap);
        }
        if (excelHead.getFieldName().equals("area")) {
            drop.setDependFieldName("city");
            Map<String,List<String>> cascadeDropDownMap = new HashMap<>();
            cascadeDropDownMap.put("成都",Arrays.asList("高新","金牛"));
            cascadeDropDownMap.put("自贡",Arrays.asList("大安"));
            cascadeDropDownMap.put("昆明",Arrays.asList("昆明1区"));
            cascadeDropDownMap.put("大理",Arrays.asList("丽江"));
            drop.setCascadeDropDownMap(cascadeDropDownMap);
        }

        if (excelHead.getFieldName().equals("unit")) {
            List<String> list = new ArrayList<>();
            //todo 数据可来自数据库
            drop.setSingleDropDownList(list);
        }

        if (excelHead.getFieldName().equals("role")) {
            drop.setSingleDropDownList(Arrays.asList("开发","测试","运维"));
        }
        return drop;
    }
}
