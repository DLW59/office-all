package com.dlw.architecture.office.template;

import lombok.extern.slf4j.Slf4j;

import java.io.*;

/**
 * @author dengliwen
 * @date 2020/6/11
 * @desc word 模板上传
 * @since 4.0.0
 */
@Slf4j
public class WordTemplateImporter extends AbstractTemplateImporter {

    /**
     * 默认上传到classpath的template/word目录，没有则创建
     */
    private static final String WORD_DEFAULT_TEMPLATE_PATH = "template/word";

    /**
     * word模板的上传  用户上传到项目指定的模板目录 eg: classpath: template/word/xx.doc
     * @param inputStream 文件流 前端上传的MultipartFile的文件流
     * @param templatePath word模板存放位置
     * @param fileName 文件名
     */
    @Override
    public void importTemplate(InputStream inputStream, String templatePath, String fileName) throws Exception {
        final File templateDirFile = getTemplateDirFile(templatePath,WORD_DEFAULT_TEMPLATE_PATH);
        doImport(inputStream,fileName,templateDirFile.getAbsolutePath());
    }
}
