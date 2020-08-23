package com.dlw.architecture.office.excel;

import com.dlw.architecture.office.model.excel.CellRange;
import com.dlw.architecture.office.model.excel.ExcelHead;
import com.dlw.architecture.office.model.excel.ExcelHeadProperty;
import lombok.Data;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.*;

/**
 * @author dengliwen
 * @date 2020/6/28
 * @desc 合并表头单元格
 * @since 4.0.0
 */
@Data
public class ExcelHeadHelper {

    /**
     * 添加要合并的单元格到sheet
     * @param sheet sheet
     * @param headProperty 表头属性对象
     */
    public static void addMergedRegionToCurrentSheet(Sheet sheet, ExcelHeadProperty headProperty) {
        final List<CellRange> cellRanges = headCellRangeList(headProperty);
        for (CellRange cellRange : cellRanges) {
            sheet.addMergedRegionUnsafe(new CellRangeAddress(cellRange.getFirstRow(),
                    cellRange.getLastRow(), cellRange.getFirstCol(), cellRange.getLastCol()));
        }
    }

    /**
     * 表头单元格合并范围算法
     * @param headProperty 表头属性
     * @return 要合并的单元格信息集合
     */
    private static List<CellRange> headCellRangeList(ExcelHeadProperty headProperty) {
        List<CellRange> cellRangeList = new ArrayList<>();
        Set<String> alreadyRangeSet = new HashSet<>();
        List<ExcelHead> headList = new ArrayList<>(headProperty.getHeadMap().values());
        for (int i = 0; i < headList.size(); i++) {
            ExcelHead head = headList.get(i);
            List<String> headNameList = head.getHeadNameList();
            for (int j = 0; j < headNameList.size(); j++) {
                if (alreadyRangeSet.contains(i + "-" + j)) {
                    continue;
                }
                alreadyRangeSet.add(i + "-" + j);
                String headName = headNameList.get(j);
                int lastCol = i;
                int lastRow = j;
                for (int k = i + 1; k < headList.size(); k++) {
                    if (headList.get(k).getHeadNameList().get(j).equals(headName)) {
                        alreadyRangeSet.add(k + "-" + j);
                        lastCol = k;
                    } else {
                        break;
                    }
                }
                Set<String> tempAlreadyRangeSet = new HashSet<>();
                outer:
                for (int k = j + 1; k < headNameList.size(); k++) {
                    for (int l = i; l <= lastCol; l++) {
                        if (headList.get(l).getHeadNameList().get(k).equals(headName)) {
                            tempAlreadyRangeSet.add(l + "-" + k);
                        } else {
                            break outer;
                        }
                    }
                    lastRow = k;
                    alreadyRangeSet.addAll(tempAlreadyRangeSet);
                }
                if (j == lastRow && i == lastCol) {
                    continue;
                }
                cellRangeList.add(new CellRange(j, lastRow, head.getColumnIndex(), headList.get(lastCol).getColumnIndex()));
            }
        }
        return cellRangeList;
    }
}
