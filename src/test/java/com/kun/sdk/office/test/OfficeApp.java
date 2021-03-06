package com.kun.sdk.office.test;

import com.dlw.architecture.office.enums.FileType;
import com.dlw.architecture.office.excel.ExcelExporter;
import com.dlw.architecture.office.excel.ExcelTemplateExporter;
import com.dlw.architecture.office.exception.OfficeException;
import com.dlw.architecture.office.model.excel.ExcelSheetModel;
import com.dlw.architecture.office.model.pdf.PdfTableModel;
import com.dlw.architecture.office.pdf.PdfExporter;
import com.dlw.architecture.office.pdf.PdfTemplateExporter;
import com.dlw.architecture.office.template.WordTemplateImporter;
import com.dlw.architecture.office.word.util.WordExporter;
import com.dlw.architecture.office.word.wrapper.WordDataWrapper;
import com.kun.sdk.office.test.example.*;
import com.kun.sdk.office.test.example.word.EasyWord;
import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

import static com.dlw.architecture.office.excel.ExcelImporter.read;


/**
 * @author dengliwen
 * @date 2020/6/15
 * @desc
 */
@SpringBootApplication
@RestController
@Slf4j
public class OfficeApp {

    public static void main(String[] args) {
        SpringApplication.run(OfficeApp.class, args);
    }

    @GetMapping("/excel/template")
    public void template(HttpServletResponse response) throws IOException, OfficeException {
        final ExcelSheetModel<Address> model1 = new ExcelSheetModel<>();
        model1.setSheetName("地址信息");
        model1.setClazz(Address.class);
        model1.setDropDownHandler(new MyDropDownHandler());
        final ExcelSheetModel model2 = ExcelSheetModel.builder()
                .sheetName("人员信息")
                .index(3)
                .clazz(Person.class)
                .build();
        final ExcelSheetModel model3 = ExcelSheetModel.builder()
                .index(2)
                .clazz(Demo.class)
                .dropDownHandler(new MyDropDownHandler())
                .build();
        ResponseUtil.initExcelResponse(response, "excel测试模板", FileType.ExcelType.XLSX);
        ExcelTemplateExporter.template(response.getOutputStream(), model1, model2, model3);
    }

    @GetMapping("/excel/export")
    public void excel(HttpServletResponse response) throws IOException, OfficeException {
        final long s = System.currentTimeMillis();
        final ExcelSheetModel<Address> model = new ExcelSheetModel<>();
        model.setSheetName("地址信息");
        List<Address> addresses = new ArrayList<>();
        Address address;
        int size = 10;
        for (int i = 0; i < size; i++) {
            address = new Address("四川", "成都", "高新");
            addresses.add(address);
        }
        model.setData(addresses);
        model.setClazz(Address.class);
        model.setDropDownHandler(new MyDropDownHandler());
        final ExcelSheetModel<Person> model2 = new ExcelSheetModel<>();
        model2.setIndex(1);
        model2.setSheetName("人员信息");
        model2.setClazz(Person.class);
        List<Person> people = new ArrayList<>();
        people.add(new Person("dlw可能是放开那可能你上课京东方难舍难分是去任务VR是无v", 27));
        people.add(new Person("jack", 22));
        model2.setData(people);
        final ExcelSheetModel model3 = ExcelSheetModel.builder()
                .index(2)
                .clazz(Demo.class)
                .dropDownHandler(new MyDropDownHandler())
                .data(Arrays.asList(new Demo()))
                .build();
        ResponseUtil.initExcelResponse(response, "excel测试数据", FileType.ExcelType.XLS);
        ExcelExporter.export(response.getOutputStream(), model, model2, model3);
        log.info("导出{}条数据耗时:{}ms", size, System.currentTimeMillis() - s);
    }

    @GetMapping("/excel/import/one")
    public String excelImport() throws Exception {
        final long s = System.currentTimeMillis();
        String path = "E:\\excel\\excel测试模板.xls";
        File file = new File(path);
        FileInputStream inputStream = new FileInputStream(file);
        final List<Address> addresses = read(inputStream, Address.class.getName());
        log.info("读取耗时:{}ms", System.currentTimeMillis() - s);
//        return "读取" + addresses.size() + "条数据耗时:" + (System.currentTimeMillis() - s) + "ms";
        return addresses.toString();
    }

    @GetMapping("/excel/import/multi")
    public String excelImports() throws Exception {
        String path = "E:\\excel\\excel测试模板.xlsx";
        File file = new File(path);
        FileInputStream inputStream = new FileInputStream(file);
        final List<Class<?>> classes = Arrays.asList(Address.class, Person.class, Demo.class);
        final List<List<?>> lists = read(inputStream, classes);
        List<Address> addresses = (List<Address>) lists.get(0);
        List<Person> people = (List<Person>) lists.get(1);
        log.info(addresses.toString());
        log.info(people.toString());
        return lists.toString();
    }


    @GetMapping("/word/export/complex")
    public void complexWord(HttpServletResponse response) throws Exception {
        final Person person = person();
        final Map<String, Object> map = WordDataWrapper.wrap(person);
        ResponseUtil.initWordResponse(response, "word导出", FileType.wordType.DOC);
        FileInputStream inputStream = new FileInputStream(new File("C:\\Users\\ASUS\\Desktop\\word.doc"));
        WordExporter.exportWord(map, inputStream, response.getOutputStream());
    }

    @GetMapping("/word/export/easy")
    public void easyWord(HttpServletResponse response) throws Exception {
        Person person = new Person();
        person.setUsername("ddd");
        final EasyWord easyWord = new EasyWord("简单word模板测试", person, 25);
        ResponseUtil.initWordResponse(response, "word导出", FileType.wordType.DOC);
        FileInputStream inputStream = new FileInputStream(new File("C:\\Users\\ASUS\\Desktop\\easyWord.doc"));
        WordExporter.exportWord(easyWord, inputStream, response.getOutputStream());
    }

    @GetMapping("/word/export/payment")
    public void payWord(HttpServletResponse response) throws Exception {
        final Payment payment = new Payment();
        final Map<String, Object> map = WordDataWrapper.wrap(payment);
        ResponseUtil.initWordResponse(response, "payment导出", FileType.wordType.DOCX);
        FileInputStream inputStream = new FileInputStream(new File("C:\\Users\\ASUS\\Desktop\\payment.docx"));
        WordExporter.exportWord(map, inputStream, response.getOutputStream());
    }


    @GetMapping("/word/template")
    @Deprecated
    public void wordTemplate() throws Exception {
        WordTemplateImporter importer = new WordTemplateImporter();
        final File file = new File("C:\\Users\\ASUS\\Desktop\\word.doc");
        FileInputStream inputStream = new FileInputStream(file);
        importer.importTemplate(inputStream, "", file.getName());
    }

    private Person person() {
        Person person = new Person();
        person.setTitle("测试word模板");
        person.setUsername("张三");
        person.setAge(22);
        person.setImg("http://deepoove.com/images/icecream.png");
        person.setHobby(Arrays.asList("王者荣耀", "敲代码", "哈哈"));
        List<Person.Plan> plans = new ArrayList<>();
        plans.add(new Person.Plan("2020.1.1", "放假休息"));
        plans.add(new Person.Plan("2020.3.1", "上班"));
        person.setPlan(plans);
        List<Address> addresses = new ArrayList<>();
        addresses.add(new Address("四川", "成都"));
        addresses.add(new Address("四川", "绵阳"));
        person.setAddress(addresses);
        return person;
    }

    @GetMapping("/pdf")
    public void pdf(HttpServletResponse response) throws Exception {
        Map<String, Object> map = new HashMap<>();
        map.put("name", "dlw");
        map.put("age", 22);
        ResponseUtil.initPdfResponse(response, "pdf模板导出测试");
        PdfTemplateExporter.exportPdfByTemplate(map, "template/pdf/test.pdf", 2, response.getOutputStream());
    }

    @GetMapping("/pdf/export")
    public void pdfExport(HttpServletResponse response) throws Exception {
        List<Address> addresses = new ArrayList<>();
        addresses.add(new Address("四川", "C:\\Users\\ASUS\\Pictures\\Saved Pictures\\1.jpg"));
        addresses.add(new Address("四川", "C:\\Users\\ASUS\\Pictures\\Saved Pictures\\1.jpg"));
        List<Person> people = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            people.add(new Person("德莱文" + i, i));
        }
        final List<PdfTableModel> models = PdfTableModel.builder()
                .of(1, Address.class, addresses)
                .of(2, Person.class, people)
                .list();
        ResponseUtil.initPdfResponse(response, "pdf导出测试");
        Thread.sleep(3000);
        PdfExporter.export(models, response.getOutputStream());
    }

}
