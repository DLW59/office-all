package com.dlw.architecture.office.template;

import java.io.InputStream;

/**
 * @author dengliwen
 * @date 2020/6/20
 * @desc 模板导入接口
 * @since 4.0.0
 */
public interface TemplateImporter {

    /**
     * 模板的上传  用户上传到项目指定的模板目录 eg: classpath: template/word/xx.doc
     * @param inputStream 文件流 前端上传的MultipartFile的文件流
     * @param templatePath word模板存放位置
     * @param fileName 文件名
     * @throws Exception
     */
    void importTemplate(InputStream inputStream, String templatePath, String fileName) throws Exception;
}
