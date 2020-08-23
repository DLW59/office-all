package com.dlw.architecture.office.pdf;

import com.dlw.architecture.office.exception.OfficeException;
import com.dlw.architecture.office.util.ConvertUtil;
import com.dlw.architecture.office.util.IoUtil;
import com.itextpdf.text.Document;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.pdf.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.util.IOUtils;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Map;

/**
 * @author dengliwen
 * @date 2020/6/16
 * @desc pdf模板导出工具
 * @since 4.0.0
 */
@Slf4j
public class PdfTemplateExporter {

    /**
     * 根据PDF模板导出数据
     *
     * @param data         填充的数据
     * @param templatePath 模板路径
     * @param pageNum      PDF模板的总页数
     * @param outputStream 输出流
     * @throws Exception
     */
    public static void exportPdfByTemplate(Object data, String templatePath, int pageNum, OutputStream outputStream) throws Exception {
        final InputStream inputStream = PdfTemplateExporter.class.getClassLoader().getResourceAsStream(templatePath);
        exportPdfByTemplate(data,inputStream,pageNum,outputStream);
    }

    /**
     * 根据PDF模板导出数据
     *
     * @param data         填充的数据
     * @param inputStream 输入流
     * @param pageNum      PDF模板的总页数
     * @param outputStream 输出流
     * @throws Exception
     */
    public static void exportPdfByTemplate(Object data, InputStream inputStream, int pageNum, OutputStream outputStream) throws Exception {
        final Map<String, Object> map = ConvertUtil.beanToMap(data);
        ByteArrayOutputStream[] bos = new ByteArrayOutputStream[pageNum];
        try {
            if (null == inputStream) {
                throw new OfficeException("pdf template does not exist");
            }
            ByteArrayOutputStream os = IoUtil.reuseInputStream(inputStream);
            BaseFont bfChinese = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H",
                    BaseFont.NOT_EMBEDDED);
            ArrayList<BaseFont> fontList = new ArrayList<>();
            fontList.add(bfChinese);
            Document document = new Document(PageSize.A4);
            PdfCopy copy = new PdfCopy(document, outputStream);
            document.open();
            for (int i = 0; i < pageNum; i++) {
                bos[i] = new ByteArrayOutputStream();
                PdfReader reader = new PdfReader(new ByteArrayInputStream(os.toByteArray()));
                PdfStamper stamper = new PdfStamper(reader, bos[i]);
                AcroFields form = stamper.getAcroFields();
                //文字类的内容处理
                form.setSubstitutionFonts(fontList);
                for (String key : form.getFields().keySet()) {
                    String value = String.valueOf(map.get(key));
                    form.setField(key, value);
                }
                stamper.setFormFlattening(true);
                stamper.close();
                reader.close();
            }

            for (int i = 0; i < pageNum; i++) {
                PdfImportedPage importPage = copy.getImportedPage(new PdfReader(bos[i].toByteArray()), i + 1);
                copy.addPage(importPage);
            }
            document.close();
        } finally {
            IOUtils.closeQuietly(outputStream);
        }
    }
}
