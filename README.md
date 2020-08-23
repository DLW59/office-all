##作者qq邮箱： 983546945@qq.com

### **office sdk 使用说明文档**
```
使用：项目添加依赖
    <dependency>
		<groupId>com.dlw.office</groupId>
		<artifactId>sdk-office</artifactId>
		<version>1.0.0</version>
	</dpendency>
```
#### 1.excel sdk使用文档

##### 1.1 excel模板下载

支持多sheet模板下载,每个sheet数据封装到ExcelSheetModel对象。

```
使用列子：
 final ExcelSheetModel<Address> model1 = new ExcelSheetModel<>();
        model1.setSheetName("地址信息");
        model1.setClazz(Address.class);
        model1.setDropDownHandler(new MyDropDownHandler());
        final ExcelSheetModel model2 = ExcelSheetModel.builder()
                .sheetName("人员信息")
                .index(1)
                .clazz(Person.class)
                .build();
        final ExcelSheetModel model3 = ExcelSheetModel.builder()
                .index(2)
                .clazz(Demo.class)
                .dropDownHandler(new MyDropDownHandler())
                .build();
        ResponseUtil.initExcelResponse(response,"excel测试模板", FileType.ExcelType.XLSX);
        ExcelTemplateExporter.template(response.getOutputStream(),model1,model2,model3);

属性：
- sheetName：设置sheet的名称
- clazz：生成模板表头对象的class
- index：用于多个sheet进行排序，升序
- dropDownHandler：用于设置下拉框数据接口，需自定义类实现此接口填充数据，支持单下拉框和级联下拉框
- data：导出时需要的数据

提供的注解：
- @ExcelSheet：用于设置sheet的属性。
    - name: 设置sheet的名称，当使用了此注解设置名称可以不设置sheetName属性，如果两个都设置了以sheetName为准。
    - headConfigPath：配置excel head的配置文件路径，多个用 ','号隔开
    - prefix：配置读取配置文件的前缀,默认是'${' ,可以自定义 
    - suffix: 读取配置文件后缀默认为'}'  
    前缀和后缀须同时配置且非空
    eg:${key} 读取excel表头, 配置文件格式：key=表头1,表头1-1
    使用：@ExcelSheet(name = "人员信息",headConfigPath = {"head1.properties", "head2.properties"},prefix = "@{",suffix = "}")
- @ExcelCell：用于设置每个单元格属性。
    - name:设置表头名称，支持从配置文件读取，默认读取方式：${key},key即为配置文件的key
    excel 表头信息须配置在工程META-INF/excel_head.properties文件中.
    - index:表头的顺序
    - width：可设置每个单元格宽度，默认为20
- @ExcelHeadStyle：设置表头样式
    - fillPatternType：颜色填充类型，默认FillPatternType.SOLID_FOREGROUND
    - fillBackgroundColor：单元格前景填充颜色  默认为IndexedColors.YELLOW  

下拉框：
需调用方自己实现此接口设置数据。
eg:
public class MyDropDownHandler implements DropDownHandler {

    @Override
    public DropDownModel dropDown(ExcelHead excelHead) {
        DropDownModel drop = new DropDownModel();
        //级联下拉
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

        //单下拉框
        if (excelHead.getFieldName().equals("age")) {
            drop.setSingleDropDownList(Arrays.asList("10","20","22","23"));
        }
        return drop;
    }
}

- dependFieldName:表示该列依赖的字段名
- cascadeDropDownMap：级联下拉数据
- singleDropDownList：单下拉框数据

下载模板提供两个api:
1. ExcelTemplateExporter.template(OutputStream outputStream, ExcelSheetModel<?>... excelSheetModel)
2. ExcelTemplateExporter.template(OutputStream outputStream, FileType.ExcelType excelType,ExcelSheetModel<?>... excelSheetModel)
```

##### 1.2 excel 数据导出
导出与模板下载基本一致，只需再加上导出数据即可。

```
导出提供的api:
1. ExcelExporter.export(OutputStream outputStream, ExcelSheetModel<?>... excelSheetModel)
2. ExcelExporter.export(OutputStream outputStream,FileType.ExcelType excelType, ExcelSheetModel<?>... excelSheetModel)
```

##### 1.3 excel 导入


```
提供的api:
1. ExcelImporter.read(InputStream inputStream,Class<T> clazz) //根据@ExcelSheet设置的名称读取指定的sheet
2. ExcelImporter.read(InputStream inputStream,Class<T> clazz,String sheetName) //根据sheetName读取指定的sheet
3. ExcelImporter.read(InputStream inputStream,String className,String sheetName) //根据sheetName读取指定的sheet
4. ExcelImporter.read(InputStream inputStream,String className) //根据@ExcelSheet设置的名称读取指定的sheet
5. ExcelImporter.read(InputStream inputStream,List<Class<?>> classes)//根据@ExcelSheet设置的名称读取指定的sheet,可读取多个sheet

```

#### 2.word sdk使用文档

##### 2.1 word导出
基于poi-tl开发，可参考[官方文档](http://deepoove.com/poi-tl/)

```
提供的注解：作用于字段
- @WordText:表示为文本类型
- @WordPicture：表示图片类型，
    - width：设置图片宽度 默认200
    - height： 设置图片高度 默认200
    - format: 设置图片格式 默认png
- @WordList: 表示为列表类型
    - format：列表的格式 默认为0
      列表格式 对应关系如下 
      0 -> FMT_BULLET //● ● ●
      1 -> FMT_DECIMAL //1. 2. 3.
      2 -> FMT_DECIMAL_PARENTHESES //1) 2) 3)
      3 -> FMT_LOWER_LETTER //a. b. c.
      4-> FMT_LOWER_ROMAN //i ⅱ ⅲ
      5 -> FMT_UPPER_LETTER //A. B. C.
      6 -> FMT_UPPER_ROMAN //I. II. III.
      
- @WordLoopTable:表示固定表格类型，适用于模板有表结构只需循环填充内容

- @WordMiniTable：表示动态表格类型，适用于没有表结构，根据自己指定表头生成表格
    - header：表头

- @Name:此注解是第三方jar提供
    - value：此属性设置字段别名，对应于模板填充的名称  eg：@Name("age") 对应 {{age}}
  word模板填充内容占位符如下：
    - 文本 {{var}}
    - 图片 {{@var}}
    - 列表 {{*var}}
    - 动态表格 {{#var}}
    - 固定表格   {{address}}省份      |   城市
                [province]           |   [city]
    - 文档 {{+var}}
    var表示需要填充内容的字段名称或字段别名，需自己定义


使用样例：
 @GetMapping("/word/export")
    public void word(HttpServletResponse response) throws Exception {
        final Person person = person();  //1
        final Map<String, Object> map = WordDataWrapper.wrap(person); //2
        ResponseUtil.initWordResponse(response,"word导出", FileType.wordType.DOC); //3
        FileInputStream inputStream = new FileInputStream(new File("C:\\Users\\ASUS\\Desktop\\word.doc")); //4
        WordExporter.exportWord(map,inputStream,response.getOutputStream()); //5
    }
1.准备导出的数据
2.包装数据，SDK会转换注解的类型为word所需类型。如果导出的类型比较单一，只有文本类型，可跳过这步。
3.指定导出word文件名称和类型 doc和docx
4.获取word模板，可以在项目classpath，也可通过网络下载。
5.导出word文件。
```


#### 3.pdf sdk使用文档

##### 3.1 pdf导出

可根据制定的PDF模板导出或通过数据动态导出
```
提供的注解：

- @PdfTable：设置PDF导出实体的属性
    - title：pdf table 标题
    - titleFontType：标题字体类型 默认加粗 PdfFontType.BOLD
    - titleSize：标题字体大小 -1表示默认值  如果不使用默认值则设置的值必须大于0
    字体大小 此属性配合titleFontType一起使用时，如果titleSize=-1表示使用titleFontType的默认size值，
    否则使用titleSize属性值设置字体大小
    - titleColor: 设置标题颜色
    - headFontType：表头字体类型 默认PdfFontType.NORMAL
    - headSize：-1表示默认值  如果不使用默认值则设置的值必须大于0
      字体大小 此属性配合headFontType一起使用时，如果headSize=-1表示使用headFontType的默认size值，
      否则使用headSize属性值设置字体大小
    - headColor：设置表头颜色
    - totalWidth：总宽度  表示表格总宽度 如果不是固定宽度模式则可以不设置
    - isLockWidth：是否固定宽度 默认不固定 自适应宽度 固定宽度需自己控制宽度 
      如果为true则totalWidth属性值必须大于0
    - isHead: 是否有表头 默认有

- @PdfCell: 设置pdf单元格样式
    - name：表示列中文名 用于表头 当设置为有表头的时候name不能为空 必须保持表头与列字段数相等
    - index：列的索引 表示表头顺序 从0开始
    - width: 列的宽度 每列占总宽度的比列
    - cellType: 单元格类型 默认文本 PdfCellType.TEXT
    - fontType: 单元格字体类型 默认普通文本 PdfFontType.NORMAL
    - size：-1表示默认值  如果不使用默认值则设置的值必须大于0
      字体大小 此属性配合fontType一起使用时，如果size=-1表示使用fontType的默认size值，
      否则使用size属性值设置字体大小
    - cellColor：设置单元格颜色

使用样例：
  根据模板导出pdf:
  @GetMapping("/pdf")
    public void pdf(HttpServletResponse response) throws Exception {
        Map<String,Object> map = new HashMap<>();
        map.put("name","dlw");
        map.put("age", 22);   //1
        ResponseUtil.initPdfResponse(response,"pdf模板导出测试");//2
        PdfTemplateExporter.exportPdfByTemplate(map, "template/pdf/test.pdf",2,response.getOutputStream());//3
    }

1.准备导出数据
2，设置导出pdf文件名
3.导出，指定模板路径或提供远程网络下载模板的流，指定模板的页数。


  动态导出pdf:
  @GetMapping("/pdf/export")
    public void pdfExport(HttpServletResponse response) throws Exception {

        List<Address> addresses = new ArrayList<>();
        List<Person> people = new ArrayList<>();
        for (int i = 0;i<10;i++) {
            people.add(new Person("德莱文" + i,i));    //1
        }

        final List<PdfTableModel> models = PdfTableModel.builder()
                .of(1,Address.class, addresses)
                .of(2,Person.class, people)
                .list();         //2

        ResponseUtil.initPdfResponse(response,"pdf导出测试");//3
        PdfExporter.export(models,response.getOutputStream());//4
    }
1.获取导出的数据
2.构造PdfTableModel数据，需指定对个PDFtable之间的顺序、导出模型的class及导出数据。
3.设置导出pdf文件名称
4.导出pdf
```




