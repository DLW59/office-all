package com.dlw.architecture.office.excel;

import com.dlw.architecture.office.enums.FileType;
import com.dlw.architecture.office.exception.OfficeException;
import com.dlw.architecture.office.model.excel.ExcelSheetModel;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;

import java.io.OutputStream;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;


/**
 * @author dengliwen
 * @date 2020/6/17
 * @desc excel导出 支持导出多个sheet，根据index的顺序导出
 * 可通过参数 {@link FileType.ExcelType} 指定导出的excel格式
 * @since 4.0.0
 */
@Slf4j
public class ExcelExporter {

	/**
	 * 导出excel数据 ，默认导出xlsx格式 支持多sheet导出
     * 此接口用于使用@ExcelColumn注解导出的采场景
	 * @param outputStream 输出流
	 * @param excelSheetModel excel数据模型
	 */
	public static void export(OutputStream outputStream, ExcelSheetModel<?>... excelSheetModel) throws OfficeException {
		export(outputStream, FileType.ExcelType.XLSX, excelSheetModel);
	}

	/**
	 * 导出excel数据 可指定导出excel格式
     * 此接口用于使用@ExcelColumn注解导出的采场景
	 * @param outputStream 输出流
	 * @param excelType excel格式
	 * @param excelSheetModel excel数据模型
	 */
	public static void export(OutputStream outputStream,FileType.ExcelType excelType, ExcelSheetModel<?>... excelSheetModel) throws OfficeException {
	    final List<ExcelSheetModel<?>> excelSheetModels = Arrays.asList(excelSheetModel);
		Workbook wb = ExcelHelper.getWorkbook(excelType);
	    try {
            excelSheetModels.sort(Comparator.comparing(ExcelSheetModel::getIndex));
            for (ExcelSheetModel<?> model : excelSheetModels) {
                ExcelHelper.createSheet(wb,model);
            }
            wb.write(outputStream);
        }catch (Exception e) {
            throw new OfficeException("failed to export excel file " + e.getMessage(),e);
        }finally {
			IOUtils.closeQuietly(wb);
		}
	}



}
