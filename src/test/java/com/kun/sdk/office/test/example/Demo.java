package com.kun.sdk.office.test.example;

import com.dlw.architecture.office.annotation.ExcelCell;
import com.dlw.architecture.office.annotation.ExcelHeadStyle;
import com.dlw.architecture.office.annotation.ExcelSheet;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.List;

/**
 * @author dengliwen
 * @date 2020/6/30
 * @desc
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@ExcelSheet(name = "测试",headConfigPath = {"head1.properties", "conf/head2.properties"})
public class Demo {
    @ExcelCell(name = {"${demo.province}"}, index = 1)
    private String province;
    @ExcelCell(name = {"${demo.city}"}, index = 2)
    private String city;
    @ExcelCell(name = {"员工基本信息详情","地址信息","区"}, index = 3)
    private String area;

    @ExcelCell(name = {"员工基本信息详情","身份信息","姓名"}, index = 4)
    private String name;
    @ExcelCell(name = {"员工基本信息详情","身份信息","年龄"}, index = 5)
    private Integer age;

    @ExcelCell(name = {"员工基本信息详情","职位"}, index = 6)
    private String role;
    @ExcelCell(name = {"员工基本信息详情","工龄"}, index = 7)
    private String year;

    @ExcelCell(name = {"家庭人员情况","人口数"})
    private String num;
    @ExcelCell(name = {"家庭人员情况","关系"})
    private String relation;
    @ExcelCell(name = {"家庭人员情况","电话"})
    private String phone;

    @ExcelCell(name = {"家庭年度总收入(万元)"},width = 30)
    @ExcelHeadStyle(fillBackgroundColor = IndexedColors.RED)
    private String total;
}
