package com.kun.sdk.office.test.example.word;

import com.deepoove.poi.el.Name;
import com.kun.sdk.office.test.example.Person;
import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * @author dengliwen
 * @date 2020/7/13
 * @desc
 */
@Data
@AllArgsConstructor
public class EasyWord {

    private String title;

    private Person person;

    @Name("ages")
    private int age;
}
