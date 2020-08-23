package com.dlw.architecture.office.excel;

import com.dlw.architecture.office.enums.FileType;
import com.dlw.architecture.office.exception.OfficeException;
import com.dlw.architecture.office.model.excel.ExcelSheetModel;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;

import java.io.OutputStream;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.Objects;

/**
 * @author dengliwen
 * @date 2020/6/17
 * @desc excel模板导出 默认导出为xlsx格式，可通过参数 {@link FileType.ExcelType} 指定导出的excel格式
 * @since 4.0.0
 */
public class ExcelTemplateExporter {

    /**
     * excel模板下载
     * @param outputStream 输出流
     * @param excelSheetModel excel模板下载数据封装模型
     */
    public static void template(OutputStream outputStream, ExcelSheetModel<?>... excelSheetModel) throws OfficeException {
        template(outputStream, FileType.ExcelType.XLSX,excelSheetModel);
    }

    /**
     * excel模板下载
     * @param outputStream 输出流
     * @param excelType excel格式
     * @param excelSheetModel excel模板下载数据封装模型
     */
    public static void template(OutputStream outputStream, FileType.ExcelType excelType, ExcelSheetModel<?>... excelSheetModel) throws OfficeException {
        final List<ExcelSheetModel<?>> excelSheetModels = Arrays.asList(excelSheetModel);
        Workbook wb = ExcelHelper.getWorkbook(excelType);
        try {
            excelSheetModels.sort(Comparator.comparing(ExcelSheetModel::getIndex));
            for (ExcelSheetModel<?> model : excelSheetModels) {
                final Class<?> clazz = model.getClazz();
                if (Objects.isNull(clazz)) {
                    continue;
                }
                ExcelHelper.createHead(wb,model);
            }
            wb.write(outputStream);
        }catch (Exception e) {
            throw new OfficeException("failed to export excel template " + e.getMessage(),e);
        }finally {
            IOUtils.closeQuietly(wb);
        }
    }

}
