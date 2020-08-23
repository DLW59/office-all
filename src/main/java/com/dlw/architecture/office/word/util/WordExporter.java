package com.dlw.architecture.office.word.util;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.dlw.architecture.office.annotation.WordLoopTable;
import com.dlw.architecture.office.exception.OfficeException;
import org.apache.poi.util.IOUtils;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;
import java.util.Objects;

/**
 * @author dengliwen
 * @date 2020/6/9
 * @desc word导出工具类 用于最终的word导出
 * @since 4.0.0
 */
public class WordExporter {

    /**
     * 导出word文件
     * @param data 导出的数据
     * @param templatePath word模板路径
     * @param outputStream 输出流
     * @throws IOException
     */
    public static void exportWord(Object data, String templatePath,
                                  OutputStream outputStream) throws Exception {
        final InputStream inputStream = WordExporter.class.getClassLoader().getResourceAsStream(templatePath);
        doExport(data,inputStream,outputStream);
    }


    /**
     * 导出word文件
     * @param data 导出的数据
     * @param inputStream 模板数据流
     * @param outputStream 输出流
     * @throws IOException
     */
    public static void exportWord(Object data, InputStream inputStream, OutputStream outputStream) throws Exception {
        doExport(data, inputStream, outputStream);
    }

    /**
     * 导出word实现
     * @param data 导出数据
     * @param inputStream 输入流
     * @param outputStream 输出流
     * @throws Exception
     */
    private static void doExport(Object data, InputStream inputStream, OutputStream outputStream) throws Exception {
        if (null == inputStream) {
            throw new OfficeException("word template does not exist");
        }
        XWPFTemplate template = null;
        try {
            if (data instanceof Map) {
                Map map = (Map) data;
                final Object o = map.get(WordLoopTable.class.getName());
                if (Objects.nonNull(o)) {
                    Configure configure = (Configure) o;
                    template = XWPFTemplate.compile(inputStream,configure).render(data);
                }
            }else {
                template = XWPFTemplate.compile(inputStream).render(data);
            }
            template.write(outputStream);
        }finally {
            IOUtils.closeQuietly(template);
        }
    }
}
