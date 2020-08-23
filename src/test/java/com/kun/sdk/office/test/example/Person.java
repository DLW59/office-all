package com.kun.sdk.office.test.example;

import com.deepoove.poi.el.Name;
import com.dlw.architecture.office.annotation.*;
import com.dlw.architecture.office.enums.ColorType;
import com.dlw.architecture.office.enums.PdfFontType;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.List;

/**
 * @author dengliwen
 * @date 2020/6/11
 * @desc
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@PdfTable(title = "人员信息:", titleFontType = PdfFontType.SECOND_TITLE)
@ExcelSheet(name = "人员信息",headConfigPath = {"head1.properties"},prefix = "@{",suffix = "}")
public class Person {

    @WordText
    @ExcelCell(name = {"标题"},index = 0)
    @ExcelHeadStyle(fillBackgroundColor = IndexedColors.BLUE)
    private String title;

    @WordText
    @ExcelCell(name = {"@{vip.username}"})
    @PdfCell(name = "姓名",index = 1,width = 1,size = 15,cellColor = ColorType.GRAY)
    private String username;

    @WordText
    @Name("ages")
    @ExcelCell(name = {"会员信息","年龄"})
    @PdfCell(name = "年龄",index = 2,width = 1,fontType = PdfFontType.UNDERLINE,cellColor = ColorType.RED)
    private int age;


    @WordPicture(width = 100,height = 100)
    @ExcelCell(name = {"会员信息","头像"})
    private String img;

    @WordList(format = 1)
    private List<String> hobby;

    @WordMiniTable(header = {"时间","打算"})
    private List<Plan> plan;

    @WordLoopTable
    private List<Address> address;

    public Person(String username, int age) {
        this.username = username;
        this.age = age;
    }

    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    public static class Plan {

        private String time;
        private String doSomething;
    }

    public Person(String username) {
        this.username = username;
    }
}
