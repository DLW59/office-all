package com.dlw.architecture.office.template;

import java.io.File;
import java.io.InputStream;

/**
 * @author dengliwen
 * @date 2020/6/16
 * @desc pdf模板上传操作类
 * @since 4.0.0
 */
public class PdfTemplateImporter extends AbstractTemplateImporter {

    /**
     * 默认上传到classpath的template/pdf目录，没有则创建
     */
    private static final String PDF_DEFAULT_TEMPLATE_PATH = "template/pdf";

    /**
     * word模板的上传  用户上传到项目指定的模板目录 eg: classpath: template/pdf/xx.pdf
     * @param inputStream 文件流
     * @param templatePath pdf模板存放位置
     * @param fileName 文件名
     */
    @Override
    public void importTemplate(InputStream inputStream, String templatePath, String fileName) throws Exception {
        final File templateDirFile = getTemplateDirFile(templatePath,PDF_DEFAULT_TEMPLATE_PATH);
        doImport(inputStream,fileName,templateDirFile.getAbsolutePath());
    }
}
