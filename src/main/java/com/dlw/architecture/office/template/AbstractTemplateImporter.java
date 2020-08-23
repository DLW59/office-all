package com.dlw.architecture.office.template;

import com.dlw.architecture.office.exception.OfficeException;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;

/**
 * @author dengliwen
 * @date 2020/6/16
 * @desc 模板上传接口  适用于模板存放在工程目录下
 * @since 4.0.0
 */
public abstract class AbstractTemplateImporter implements TemplateImporter {

    /**
     * 默认下载到classpath下
     */
    private static final String TEMPLATE_PATH_PREFIX = AbstractTemplateImporter.class.getClassLoader().getResource("").getPath();


    /**
     * 生成模板存放文件夹
     * @param templatePath 指定的模板存放路径
     * @param defaultTemplatePath 默认存放模板路径
     * @return
     */
    protected File getTemplateDirFile(String templatePath, String defaultTemplatePath) {
        if (StringUtils.isBlank(templatePath)) {
            templatePath = defaultTemplatePath;
        }
        File fileDir = new File(TEMPLATE_PATH_PREFIX + templatePath);
        if(!fileDir.exists()){
            // 递归生成文件夹
            fileDir.mkdirs();
        }
        return fileDir;
    }

    /**
     * 上传模板
     * @param inputStream 文件流
     * @param fileName 模板文件名称
     * @param absolutePath 模板文件夹路径
     * @throws Exception
     */
    protected void doImport(InputStream inputStream, String fileName, String absolutePath) throws Exception {
        FileOutputStream os = null;
        try {
            File file = new File(absolutePath + File.separator + fileName);
            if (file.exists()) {
                file.delete();
            }
            os = new FileOutputStream(file);
            int bytesRead;
            byte[] buffer = new byte[8192];
            while ((bytesRead = inputStream.read(buffer, 0, 8192)) != -1) {
                os.write(buffer, 0, bytesRead);
            }
        }catch (Exception e) {
            throw new OfficeException("The upload word template is error:" + e.getMessage(),e);
        }finally {
            if (os != null) {
                os.close();
            }
            inputStream.close();
        }
    }
}
