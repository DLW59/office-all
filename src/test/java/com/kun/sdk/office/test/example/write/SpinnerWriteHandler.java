package com.kun.sdk.office.test.example.write;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

import java.util.Map;

/**
 * @author dengliwen
 * @date 2020/6/24
 * @desc
 */
@NoArgsConstructor
public class SpinnerWriteHandler implements SheetWriteHandler {

    @Setter
    @Getter
    private Map<Integer,String[]> dropDownMap;

    public SpinnerWriteHandler(Map<Integer, String[]> dropDownMap) {
        this.dropDownMap = dropDownMap;
    }

    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {

    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
//定义一个map key是需要添加下拉框的列的index value是下拉框数据
        //获取工作簿
        Sheet sheet = writeSheetHolder.getSheet();
        ///开始设置下拉框
        DataValidationHelper helper = sheet.getDataValidationHelper();
        //设置下拉框
        for (Map.Entry<Integer, String[]> entry : dropDownMap.entrySet()) {
            /*起始行、终止行、起始列、终止列  起始行为1即表示表头不设置**/
            CellRangeAddressList addressList = new CellRangeAddressList(1, 65535, entry.getKey(), entry.getKey());
            /*设置下拉框数据**/
            DataValidationConstraint constraint = helper.createExplicitListConstraint(entry.getValue());
            DataValidation dataValidation = helper.createValidation(constraint, addressList);
            /*处理Excel兼容性问题**/
            if (dataValidation instanceof XSSFDataValidation) {
                dataValidation.setSuppressDropDownArrow(true);
                dataValidation.setShowErrorBox(true);
            } else {
                dataValidation.setSuppressDropDownArrow(false);
            }
            sheet.addValidationData(dataValidation);
        }
    }
}
