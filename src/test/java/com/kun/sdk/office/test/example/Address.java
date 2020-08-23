package com.kun.sdk.office.test.example;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.dlw.architecture.office.annotation.ExcelCell;
import com.dlw.architecture.office.annotation.ExcelSheet;
import com.dlw.architecture.office.annotation.PdfCell;
import com.dlw.architecture.office.annotation.PdfTable;
import com.dlw.architecture.office.enums.ColorType;
import com.dlw.architecture.office.enums.PdfFontType;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @author dengliwen
 * @date 2020/6/12
 * @desc
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@PdfTable(title = "地址信息:",titleColor = ColorType.ORANGE,titleSize = 20)
@ExcelSheet(name = "地址信息")
public class Address {

    @ExcelProperty(value = {"标题1","标题1-1"},index = 0)
    @ColumnWidth(value = 20)
    @ExcelCell(name = {"省"}, index = 1,width = 25)
    @PdfCell(name = "省份",index = 1,width = 0.1f,fontType = PdfFontType.BOLD,size = 20)
    private String province;
    @ExcelProperty(value = {"标题1","标题1-2"},index = 1)
    @ColumnWidth(value = 20)
    @ExcelCell(name = {"市"}, index = 2)
    @PdfCell(name = "城市",index = 2,width = 0.1f,fontType = PdfFontType.BOLD)
    private String city;
    @ExcelProperty(value = {"标题1","标题2-1","标题2-1-1"},index = 3)
    @ColumnWidth(value = 20)
    @ExcelCell(name = {"区"}, index = 3,width = 25)
    private String area;
    @ExcelProperty(value = {"标题1"},index = 2)
    @ColumnWidth(value = 20)
    @ExcelCell(name = {"街道"}, index = 4)
    private String unit;

    public Address(String province, String city) {
        this.province = province;
        this.city = city;
    }

    public Address(String province, String city, String area) {
        this.province = province;
        this.city = city;
        this.area = area;
    }
}
