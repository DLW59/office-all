package com.dlw.architecture.office.util;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

/**
 * @author dengliwen
 * @date 2020/6/22
 * @desc io流工具类
 * @since 4.0.0
 */
public class IoUtil {

    /**
     * 重用文件流 避免输入流只能使用一次出现 stream closed 异常
     * @param inputStream 文件数据输入流
     * @return 输出流
     */
    public static ByteArrayOutputStream reuseInputStream(InputStream inputStream) throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024 * 4];
        int len;
        while ((len = inputStream.read(buffer)) > -1 ) {
            outputStream.write(buffer, 0, len);
        }
        outputStream.flush();
        return outputStream;
    }
}
