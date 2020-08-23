package com.kun.sdk.office.test.example.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.kun.sdk.office.test.example.Address;
import com.kun.sdk.office.test.example.Person;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author dengliwen
 * @date 2020/6/24
 * @desc
 */
@Slf4j
public class PersonListener extends AnalysisEventListener<Person> {

    @Getter
    private List<Person> people = new ArrayList<>();

    @Override
    public void invoke(Person person, AnalysisContext analysisContext) {

        people.add(person);
        //todo
        // 可以进行批量操作 比如每100条数据入一次库
        // opt db

}

    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        log.info("表头:{}",headMap);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {

    }
}
