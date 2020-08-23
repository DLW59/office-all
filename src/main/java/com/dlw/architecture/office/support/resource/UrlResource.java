package com.dlw.architecture.office.support.resource;


import com.dlw.architecture.office.exception.OfficeException;

import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;

/**
 * @author dengliwen
 * @date 2020/7/6
 * @desc url资源
 * @since 4.0.0
 */
public class UrlResource implements Resource {

    private final URL url;

    public UrlResource(URL url) {
        this.url = url;
    }

    @Override
    public InputStream getInputStream() throws OfficeException {
        try {
            URLConnection urlConnection = url.openConnection();
            urlConnection.connect();
            return urlConnection.getInputStream();
        } catch (IOException e) {
            e.printStackTrace();
            throw new OfficeException("Failed to read the configuration file stream",e);
        }
    }
}
