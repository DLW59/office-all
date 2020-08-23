package com.dlw.architecture.office.excel;

import com.dlw.architecture.office.annotation.ExcelCell;
import com.dlw.architecture.office.annotation.ExcelSheet;
import com.dlw.architecture.office.enums.FileType;
import com.dlw.architecture.office.excel.style.ExcelStyleHelper;
import com.dlw.architecture.office.exception.OfficeException;
import com.dlw.architecture.office.model.excel.DropDownModel;
import com.dlw.architecture.office.model.excel.ExcelHead;
import com.dlw.architecture.office.model.excel.ExcelHeadProperty;
import com.dlw.architecture.office.model.excel.ExcelSheetModel;
import com.dlw.architecture.office.util.ClassUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.lang.reflect.Field;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.*;


/**
 * @author dengliwen
 * @date 2020/6/15
 * @desc excel 操作辅助工具类
 * @since 4.0.0
 */
@Slf4j
public class ExcelHelper {

    /**
     * 校验并根据sheet名称获取sheet
     *
     * @param wb
     * @param sheetName sheet名称
     * @return
     */
    public static Sheet checkAndGetSheet(Workbook wb, String sheetName) throws OfficeException {
        Sheet sheet = wb.getSheet(sheetName);
        if (sheet == null) {
            throw new OfficeException(String.format("The sheet named %s does not exist", sheetName));
        }
        // 移除所有批注
        removeAllComment(sheet);
        return sheet;
    }

    /**
     * 移除sheet所有批注
     *
     * @param sheet
     */
    private static void removeAllComment(Sheet sheet) {
        if (sheet instanceof XSSFSheet && !((XSSFSheet) sheet).hasComments()) {
            return;
        }
        try {
            Map<CellAddress, ? extends Comment> cellCommentMap = sheet.getCellComments();
            if (cellCommentMap == null || cellCommentMap.isEmpty()) {
                return;
            }
            cellCommentMap.forEach((cellAddress, comment) -> {
                Row row = sheet.getRow(cellAddress.getRow());
                if (row != null) {
                    Cell cell = row.getCell(cellAddress.getColumn());
                    if (cell != null) {
                        cell.removeCellComment();
                    }
                }
            });
        } catch (Exception e) {
            log.error("移除所有批注异常");
            e.printStackTrace();
        }
    }

    /**
     * 单sheet页excel读取多少行到多少行的数据
     *
     * @param sheet
     * @param startRowIndex 开始行数
     * @param endRowIndex   结束行数
     *                      <p>
     *                      startRowIndex = 0 endRowIndex = 1 表示读取表头
     * @return
     */
    public static List<String[]> readSheetRowData(Sheet sheet, int startRowIndex, int endRowIndex) {
        List<String[]> rowList = new ArrayList<>();
        try {
            Row row;
            for (int i = startRowIndex; i <= endRowIndex; i++) {
                row = sheet.getRow(i);
                if (row == null) {
                    continue;
                }
                int len = row.getPhysicalNumberOfCells();
                String[] rowArr = new String[len];
                for (int k = 0; k < len; k++) {
                    rowArr[k] = getCellValue(row.getCell(k));
                }
                rowList.add(rowArr);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return rowList;
    }

    /**
     * 获取单元格值
     * @param cell 单元格
     * @return
     */
    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        try {
            final CellType cellType = cell.getCellType();
            switch (cellType) {
                case BOOLEAN:
                    // 得到Boolean对象的方法
                    return cell.getBooleanCellValue() + "";
                case NUMERIC:
                    // 先看是否是日期格式
                    if (DateUtil.isCellDateFormatted(cell)) {
                        // 读取日期格式
                        return cell.getDateCellValue() + "";
                    } else {
                        // 读取数字
                        // value = cell.getNumericCellValue() + "";
                        // 数字强转字符串
                        DataFormatter dataFormatter = new DataFormatter();
                        return dataFormatter.formatCellValue(cell);
                    }
                case FORMULA:
                    // 读取公式
                    return cell.getStringCellValue();
                case STRING:
                    // 读取String
                    return cell.getRichStringCellValue().toString() + "";
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "";
    }

    /**
     * 获取sheet数据
     *
     * @param sheet
     * @param headIndex 表头下标索引
     * @return
     */
    public static List<String[]> readSheetData(Sheet sheet, int headIndex) {
        Row row0 = sheet.getRow(headIndex);
        int len = row0.getPhysicalNumberOfCells();
        List<String[]> resultList = new ArrayList<>();
        // 读取结果集
        Iterator<Row> rowIterator = sheet.rowIterator();
        Iterator<Cell> cellIterator;
        Row row;
        Cell cell;
        while (rowIterator.hasNext()) {
            row = rowIterator.next();
            if (row == null) {
                continue;
            }
            int rowIndex = row.getRowNum();
            if (rowIndex <= headIndex) {
                continue;
            }
            // 记录读取信息,多加1列,记录行号
            String[] str = new String[len + 1];
            cellIterator = row.cellIterator();
            boolean isNotBlank = false;
            int colIndex;
            while (cellIterator.hasNext()) {
                cell = cellIterator.next();
                if (cell == null) {
                    continue;
                }
                colIndex = cell.getColumnIndex();
                if (colIndex >= len) {
                    break;
                }
                str[colIndex] = getCellValue(cell);
                if (!isNotBlank && StringUtils.isNotBlank(str[colIndex])) {
                    isNotBlank = true;
                }
            }
            if (isNotBlank) {
                str[len] = Integer.toString(rowIndex + 1); // 记录行号
                resultList.add(str);
            }
        }
        return resultList;
    }

    /**
     * 创建表头
     *
     * @param wb    工作簿
     * @param model excel属性模型
     */
    public static ExcelHeadProperty createHead(Workbook wb, ExcelSheetModel<?> model) throws OfficeException {
        if (model.getClazz() == null) {
            throw new OfficeException("head class is null");
        }
        final ExcelHeadProperty headProperty = ExcelHelper.getHeaders(wb,model.getClazz());
        final Map<Integer, Integer> headerWidthMap = ExcelHelper.getHeaderWidth(model.getClazz());
        final int headRowNumber = headProperty.getHeadRowNumber();
        Sheet sheet;
        if (StringUtils.isNotBlank(model.getSheetName())) {
            sheet = wb.createSheet(model.getSheetName());
        }else {
            final String sheetName = validateClass(model.getClazz());
            sheet = wb.createSheet(sheetName);
        }
        final Map<Integer, CellStyle> headStyleMap = headProperty.getHeadStyleMap();
        ExcelHeadHelper.addMergedRegionToCurrentSheet(sheet, headProperty);
        for (int i = 0; i < headRowNumber; i++) {
            Row row = sheet.createRow(i);
            for (Map.Entry<Integer, ExcelHead> entry : headProperty.getHeadMap().entrySet()) {
                Cell cell = row.createCell(entry.getKey());
                cell.setCellValue(entry.getValue().getHeadNameList().get(i));
                cell.setCellStyle(headStyleMap.get(entry.getKey()));
                sheet.setColumnWidth(entry.getKey(), 256 * headerWidthMap.get(entry.getKey()));
            }
        }

        createDropDown(wb,model, sheet, headProperty);
        return headProperty;
    }

    /**
     * 校验class信息
     * @param clazz excel模型class
     * @param <T>
     * @return sheet名称
     * @throws OfficeException
     */
    public static <T> String validateClass(Class<T> clazz) throws OfficeException {
        final ExcelSheet excelSheet = clazz.getAnnotation(ExcelSheet.class);
        if (null == excelSheet) {
            throw new OfficeException("ExcelSheet annotation is not setting");
        }
        final String sheetName = excelSheet.name();
        if (StringUtils.isBlank(sheetName)) {
            throw new OfficeException("sheet name cannot be empty");
        }
        return sheetName;
    }

    /**
     * 创建下拉
     *
     * @param wb
     * @param model        excel模型
     * @param sheet        sheet
     * @param headProperty 表头属性
     */
    private static void createDropDown(Workbook wb, ExcelSheetModel<?> model, Sheet sheet, ExcelHeadProperty headProperty) {
        if (null == model.getDropDownHandler()) {
            return;
        }
        final Map<Integer, ExcelHead> headMap = headProperty.getHeadMap();
        Map<String, Integer> fieldIndexMap = new HashMap<>();
        for (Map.Entry<Integer, ExcelHead> entry : headMap.entrySet()) {
            fieldIndexMap.put(entry.getValue().getFieldName(),entry.getKey());
        }
        for (Map.Entry<Integer, ExcelHead> entry : headMap.entrySet()) {
            final DropDownModel dropDownModel = model.getDropDownHandler().dropDown(entry.getValue());
            if (null == dropDownModel) {
                return;
            }
            final List<String> singleDropDownList = dropDownModel.getSingleDropDownList();
            final Map<String, List<String>> cascadeDropDownMap = dropDownModel.getCascadeDropDownMap();
            final String dependFieldName = dropDownModel.getDependFieldName();
            //级联下拉
            if (StringUtils.isNotBlank(dependFieldName) && MapUtils.isNotEmpty(cascadeDropDownMap)) {
                String relateCode = dependFieldName + "_" + entry.getValue().getFieldName();
                createHideSheet(wb, cascadeDropDownMap, "hideSheet", relateCode);
                creatRelateRow(sheet, relateCode, headProperty.getHeadRowNumber(), fieldIndexMap.get(dependFieldName), entry.getKey());
            }
            //单下拉
            else if (CollectionUtils.isNotEmpty(singleDropDownList)) {
                createSingleDropDown(wb,sheet, singleDropDownList, headProperty.getHeadRowNumber(), entry.getKey());
            }
        }
    }

    /**
     * 创建关联行
     * @param sheet
     * @param relateCode 关联码
     * @param naturalRowIndex 行号
     * @param relaColIndex 依赖的列号
     * @param cascadeColIndex  作用的列号
     */
    public static void creatRelateRow(Sheet sheet, String relateCode, int naturalRowIndex, int relaColIndex,
                                      int cascadeColIndex) {
        // 得到验证对象
        DataValidation data_validation_list1 = getDataValidationByFormula(sheet, dealNameName(relateCode),
                naturalRowIndex, relaColIndex);
        String excelColumnIndex = getCellColumnFlag(relaColIndex + 1);
        int offset = naturalRowIndex;
        // HSSFSheet类型级联下拉会有问题
        if (sheet instanceof SXSSFSheet) {
            offset++;
        }
        DataValidation data_validation_list2 = getDataValidationByFormula(sheet,
                "INDIRECT(SUBSTITUTE(CONCATENATE(\"_" + relateCode + "_\",$" + excelColumnIndex + offset
                        + "),\"-\",\".\"))",
                naturalRowIndex, cascadeColIndex);
        // 工作表添加验证数据
        sheet.addValidationData(data_validation_list1);
        sheet.addValidationData(data_validation_list2);
    }

    /**
     * 获取数据验证器
     * @param sheet sheet
     * @param formulaString 验证的格式
     * @param naturalRowIndex 行号
     * @param naturalColumnIndex 列号
     * @return
     */
    private static DataValidation getDataValidationByFormula(Sheet sheet, String formulaString, int naturalRowIndex,
                                                             int naturalColumnIndex) {
        // 加载下拉列表内容
        DataValidationConstraint constraint = sheet.getDataValidationHelper()
                .createFormulaListConstraint(formulaString);
        // 设置数据有效性加载在哪个单元格上。
        // 四个参数分别是：起始行、终止行、起始列、终止列
        int firstRow = naturalRowIndex;
        int lastRow;
        if (sheet instanceof HSSFSheet) {
            lastRow = 65535;
        }else {
            lastRow = 1000000;
        }
        int firstCol = naturalColumnIndex;
        int lastCol = naturalColumnIndex;
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        // 数据有效性对象
        DataValidation data_validation_list = sheet.getDataValidationHelper().createValidation(constraint, regions);
        return data_validation_list;
    }

    /**
     * 创建单下拉框
     *
     * @param wb
     * @param sheet              sheet
     * @param singleDropDownList 单下拉框数据
     * @param headRowNumber      下拉起始行号
     * @param key                列下标
     */
    private static void createSingleDropDown(Workbook wb, Sheet sheet, List<String> singleDropDownList, int headRowNumber, Integer key) {
        //开始设置下拉框
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList addressList;
        if (sheet instanceof HSSFSheet) {
            addressList = new CellRangeAddressList(headRowNumber, 65535, key, key);
        }else {
            addressList = new CellRangeAddressList(headRowNumber, 1000000, key, key);
        }
        //fixme 2003版本下拉框超过255字符会报错,2007及以上超过255不会报错但下拉框没有值
        String formulaId = "form" + UUID.randomUUID().toString().replace("-", "");
        String hiddenSheetName = "singleSheet";
        if (wb.getSheet(hiddenSheetName) != null) {
            hiddenSheetName += 1;
        }
        //下拉框值隐藏sheet
        final Sheet hiddenSheet = wb.createSheet(hiddenSheetName);
        wb.setSheetHidden(wb.getSheetIndex(hiddenSheet),true);
        Row row;
        Cell cell;
        for (int i = 0; i < singleDropDownList.size(); i++) {
            row = hiddenSheet.createRow(i);
            //隐藏表的数据列必须和添加下拉菜单的列序号相同，否则不能显示下拉菜单
            cell = row.createCell(key);
            cell.setCellValue(singleDropDownList.get(i));
        }
        //创建"名称"标签，用于链接
        Name namedCell = wb.createName();
        namedCell.setNameName(formulaId);
        namedCell.setRefersToFormula(hiddenSheetName + "!A$1:A$" + singleDropDownList.size());

        /*设置下拉框数据**/
//        DataValidationConstraint constraint = helper.createExplicitListConstraint(singleDropDownList.toArray(new String[0]));
        final DataValidationConstraint formulaListConstraint = helper.createFormulaListConstraint(formulaId);
        DataValidation dataValidation = helper.createValidation(formulaListConstraint, addressList);
        /*处理Excel兼容性问题**/
        if (dataValidation instanceof XSSFDataValidation) {
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        } else {
            dataValidation.setSuppressDropDownArrow(false);
        }

        sheet.addValidationData(dataValidation);
    }

    /**
     * 创建excel 内容
     *
     * @param sheet  sheet
     * @param list     数据
     * @param headers  表头
     * @param rowIndex 数据起始行
     */
    public static <T> void createCont(Sheet sheet, List<T> list, String[] headers, int rowIndex) throws Exception {
        CellStyle cellStyle = ExcelStyleHelper.createDefaultContentStyle(sheet.getWorkbook());
        String[] values;
        for (int i = 0; i < list.size(); i++) {
            // Row 行,Cell 方格 , Row 和 Cell 都是从0开始计数的
            // 创建一行，在页sheet上
            Row row = sheet.createRow(rowIndex + i);
            values = getValues(list.get(i));
            for (int j = 0; j < headers.length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(values[j]);
                cell.setCellStyle(cellStyle);
            }
        }
    }

    /**
     * 获取一行的值
     *
     * @param t   行数据对象
     * @param <T> 泛型参数
     * @return 字段对应的值
     * @throws Exception
     */
    private static <T> String[] getValues(T t) throws Exception {
        Field[] fields = getAnnotationSortFields(t.getClass());
        Field field;
        List<String> valueList = new ArrayList<>();
        for (int i = 0; i < fields.length; i++) {
            field = fields[i];
            field.setAccessible(true);
            Class<?> type = field.getType();
            Object value = field.get(t);
            if (value == null) {
                valueList.add("");
            } else if (type == Date.class || type == Timestamp.class) {
                Date date = type == Date.class ? (Date) value : new Date(((Timestamp) value).getTime());
                valueList.add(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(date));
            } else {
                valueList.add(String.valueOf(value));
            }
        }
        return valueList.toArray(new String[0]);
    }

    /**
     * 获取excel表头
     *
     * @param wb
     * @param clazz 实体对象class
     * @return excel头属性
     */
    public static ExcelHeadProperty getHeaders(Workbook wb, Class clazz) throws OfficeException {
        ExcelHeadProperty headProperty = new ExcelHeadProperty(clazz);
        headProperty.initHead(wb);
        return headProperty;
    }

    /**
     * 获取excel表头对应的宽度
     *
     * @param clazz 实体对象class
     * @return 表头对应的宽度映射
     */
    public static Map<Integer, Integer> getHeaderWidth(Class<?> clazz) throws OfficeException {
        Map<Integer, Integer> headerWidthMap = new HashMap<>();
        int i = 0;
        for (Field field : getAnnotationSortFields(clazz)) {
            final Integer width = field.getAnnotation(ExcelCell.class).width();
            if (width <= 0) {
                throw new OfficeException("excel cell width cannot be less than 0");
            }
            headerWidthMap.put(i, width);
            i++;
        }
        return headerWidthMap;
    }

    /**
     * 获取带注解且排序字段
     *
     * @param aClass excel模型class
     * @return 返回特定注解且排序的字段
     */
    public static Field[] getAnnotationSortFields(Class aClass) throws OfficeException {
        List<Field> fields = Arrays.asList(ClassUtil.getAnnotationFields(aClass, ExcelCell.class));
        for (Field field : fields) {
            if (field.getAnnotation(ExcelCell.class).index() < 0) {
                throw new OfficeException("@ExcelColumn annotated index attribute value cannot be less than 0");
            }
        }
        fields.sort((o1, o2) -> {
            final int index1 = o1.getAnnotation(ExcelCell.class).index();
            final int index2 = o2.getAnnotation(ExcelCell.class).index();
            return index1 - index2;
        });
        return fields.toArray(new Field[0]);
    }

    /**
     * 创建sheet
     *
     * @param wb    工作簿
     * @param model excel所需数据封装模型
     * @param <T>   泛型参数
     * @throws Exception
     */
    public static <T> void createSheet(Workbook wb, ExcelSheetModel<T> model) throws Exception {
        final ExcelHeadProperty headProperty = createHead(wb, model);
        String sheetName = model.getSheetName();
        Sheet sheet;
        if (StringUtils.isNotBlank(sheetName)) {
            sheet = wb.getSheet(sheetName);
        }else {
            sheetName = validateClass(model.getClazz());
            sheet = wb.getSheet(sheetName);
        }
        final List<T> data = model.getData();
        if (CollectionUtils.isEmpty(data)) {
            return;
        }
        List<String> head = new ArrayList<>();
        for (Integer key : headProperty.getHeadMap().keySet()) {
            head.add(String.valueOf(key));
        }
        createCont(sheet, data, head.toArray(new String[0]), headProperty.getHeadRowNumber());
    }

    /**
     * 根据excel类型创建workbook
     *
     * @param excelType excel类型 xls xlsx
     * @return workbook
     */
    public static Workbook getWorkbook(FileType.ExcelType excelType) throws OfficeException {
        if (null == excelType) {
            throw new OfficeException("No excel type specified");
        }
        Workbook wb = null;
        switch (excelType) {
            case XLS:
                wb = new HSSFWorkbook();
                break;
            case XLSX:
                wb = new SXSSFWorkbook(1000);
                break;
        }
        return wb;
    }


    /**
     * 创建隐藏sheet
     * @param wb 工作簿
     * @param relMap 依赖级联数据
     * @param hideSheetName 隐藏sheet名称
     * @param nameName
     */
    public static void createHideSheet(Workbook wb, Map<String, List<String>> relMap, String hideSheetName,
                                       String nameName) {

        if (wb.getSheet(hideSheetName) != null) {
            hideSheetName += "1";
            createHideSheet(wb, relMap, hideSheetName, nameName);
        }
        // 名称管理器已经存在时，不创建
        if (wb.getName(dealNameName(nameName)) == null) {
            // 创建一个隐藏页和隐藏数据集
            creatHideSheet(wb, relMap, hideSheetName);
            // 设置名称集数据
            creatExcelNameList(wb, relMap, hideSheetName, nameName);
        }
    }

    /**
     * 处理名称管理名称不合法情况
     * 前面统一加上下划线，由于级联关系时存在数字开头的值无法选择级联关系（虽然通过函数判断是否数字开头，但可能由于函数嵌套过多，excel中没能读取到）
     * 名称管理不能使用下划线(_)、点号(.)和反斜杠(/)以外的其他符号，可以使用?，但不能以?开头
     *
     * @param name
     * @return
     */
    private static String dealNameName(String name) {
        name = "_" + name;
        if (name.indexOf(' ') != -1) {
            name = name.replace(" ", "_");
        }
        if (name.contains("-")) {
            name = name.replaceAll("-", "."); // 方便替换回来
        }
        return name;
    }

    /**
     * 创建数据域（下拉联动的数据）
     *
     * @param workbook
     * @param hideSheetName
     *            数据域名称
     */
    private static void creatHideSheet(Workbook workbook, Map<String, List<String>> relaMap,
                                       String hideSheetName) {
        Sheet sheet = workbook.createSheet(hideSheetName);
        // 隐藏sheet
        workbook.setSheetHidden(workbook.getSheetIndex(sheet), true);
        int rowRecord = 0;
        // 名称管理器
        Row row = sheet.createRow(rowRecord);
        creatAssitRow(row, relaMap.keySet());
        rowRecord++;
        for (Map.Entry<String, List<String>> entry : relaMap.entrySet()) {
            List<String> list = new ArrayList<>();
            list.add(0, entry.getKey());
            if (entry.getValue() == null) {
                continue;
            }
            list.addAll(entry.getValue());
            creatRow(sheet, rowRecord, list);
            rowRecord++;
        }
    }

    /**
     * 创建行
     * @param row 当前行
     * @param set 下拉数据key集合
     */
    public static void creatAssitRow(Row row, Set<String> set) {
        if (set != null) {
            int i = 0;
            for (String cellValue : set) {
                // 注意列是从（1）下标开始
                Cell cell = row.createCell(i++);
                cell.setCellValue(cellValue);
            }
        }
    }

    /**
     * 创建行数据
     * @param sheet sheet
     * @param rowRecord 行号
     * @param list 每个下拉选项值及对应的子数据
     */
    private static void creatRow(Sheet sheet, int rowRecord, List<String> list) {
        if (sheet == null) {
            return;
        }
        Row row = sheet.createRow(rowRecord);
        if (CollectionUtils.isNotEmpty(list)) {
            int i = 0;
            for (String cellValue : list) {

                Cell cell = row.createCell(i++);
                try {
                    cell.setCellValue(cellValue);
                } catch( IllegalArgumentException e) {
                    cell.setCellValue("字数超长，无法显示！");
                }
            }
        }
    }

    /**
     * 名称管理
     *
     * @param workbook
     * @param hideSheetName
     *            数据域的sheet名
     */
    private static void creatExcelNameList(Workbook workbook, Map<String, List<String>> relaMap,
                                           String hideSheetName, String relateCode) {
        Name name;
        name = workbook.createName();
        name.setNameName(dealNameName(relateCode));
        name.setRefersToFormula(hideSheetName + "!$A$1:$" + getCellColumnFlag(relaMap.keySet().size()) + "$1");
        int i = 0;
        for (Map.Entry<String, List<String>> entry : relaMap.entrySet()) {
            List<String> num = new ArrayList<>();
            name = workbook.createName();
            String nameName = entry.getKey();
            if (StringUtils.isBlank(nameName)) {
                continue;
            }
            nameName = relateCode + "_" + nameName;
            if (workbook.getName(dealNameName(nameName)) == null) {
                num.add(0, nameName);
                if (entry.getValue() == null) {
                    continue;
                }
                num.addAll(entry.getValue());
                name.setNameName(dealNameName(nameName));
                name.setRefersToFormula(
                        hideSheetName + "!$B$" + (i + 2) + ":$" + getCellColumnFlag(num.size()) + "$" + (i + 2));
            }
            i++;
        }
    }

    /**
     * 获取单元格列号对应的字母
     * @param num 列号
     * @return
     */
    private static String getCellColumnFlag(int num) {
        String columFiled = "";
        if (num >= 1 && num <= 26) {
            columFiled = doHandle(num);
        } else {
            int chuNum = num / 26;
            int yuNum = num % 26;
            if (yuNum == 0) {
                chuNum = chuNum - 1;
            }
            columFiled += doHandle(chuNum);
            columFiled += doHandle(yuNum);
        }
        return columFiled;
    }

    /**
     * 处理excel单元格字母
     * @param num 列号
     * @return
     */
    private static String doHandle(final int num) {
        String[] charArr = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R",
                "S", "T", "U", "V", "W", "X", "Y", "Z" };
        if (num == 0) {
            return "Z";
        }
        return charArr[num - 1];
    }
}
