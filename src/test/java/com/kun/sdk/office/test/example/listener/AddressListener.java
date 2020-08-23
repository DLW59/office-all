package com.kun.sdk.office.test.example.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.kun.sdk.office.test.example.Address;
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
public class AddressListener extends AnalysisEventListener<Address> {

    @Getter
    private List<Address> addresses = new ArrayList<>();

    @Override
    public void invoke(Address address, AnalysisContext analysisContext) {

        addresses.add(address);
    }

    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        log.info("表头:{}",headMap);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {

    }
}
