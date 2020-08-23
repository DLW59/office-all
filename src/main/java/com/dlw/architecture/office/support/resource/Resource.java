package com.dlw.architecture.office.support.resource;


import com.dlw.architecture.office.exception.OfficeException;

import java.io.IOException;
import java.io.InputStream;

/**
 * @author dengliwen
 * @date 2020/7/6
 * @desc
 */
public interface Resource {

    /**
     * 获取资源的流数据
     * @return 输入流
     * @throws OfficeException
     */
    InputStream getInputStream() throws OfficeException;
}
