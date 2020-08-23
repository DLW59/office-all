package com.dlw.architecture.office.pdf;

import com.dlw.architecture.office.exception.OfficeException;
import com.dlw.architecture.office.model.pdf.PdfTableModel;
import com.dlw.architecture.office.pdf.style.PdfStyle;
import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.util.IOUtils;

import java.io.OutputStream;
import java.util.Collection;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @author dengliwen
 * @date 2020/6/18
 * @desc pdf 导出工具类
 * @since 4.0.0
 */
public class PdfExporter {

    /**
     * 根据注解导出数据 可导出多个PdfPTable
     *
     * @param models  PDF对象数据List
     */
    public static void export(List<PdfTableModel> models, OutputStream outputStream) throws OfficeException {
        if (CollectionUtils.isEmpty(models)) {
            throw new OfficeException("export data is empty");
        }
        models = models.stream().filter(pdfTableModel -> CollectionUtils.isNotEmpty(pdfTableModel.getData()))
                .sorted(Comparator.comparingInt(PdfTableModel::getIndex))
                .collect(Collectors.toList());
        try {
            Document document = PdfStyle.getDocument();
            PdfWriter.getInstance(document, outputStream);
            document.open();
            for (PdfTableModel tableModel : models) {
                final Collection<?> collection = tableModel.getData();
                final PdfPTable table = PdfUtil.createTable(tableModel.getClazz(), collection);
                if (null == table) {
                    continue;
                }
                document.add(table);
            }
            document.close();
        } catch (Exception e) {
            throw new OfficeException("failed to create PDF document: " + e.getMessage(), e);
        }finally {
            IOUtils.closeQuietly(outputStream);
        }
    }
}
