package com.dlw.architecture.office.support.loader;

import com.dlw.architecture.office.exception.OfficeException;
import com.dlw.architecture.office.support.resource.Resource;
import com.dlw.architecture.office.support.resource.ResourceLoader;
import lombok.Getter;
import org.apache.poi.util.IOUtils;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Properties;

/**
 * @author dengliwen
 * @date 2020/7/5
 * @desc 抽象属性解析器
 * @since 4.0.0
 */
@Getter
public abstract class AbstractPropertyLoader implements PropertyLoader {

    private ResourceLoader resourceLoader;

    private Properties properties;

    public AbstractPropertyLoader(ResourceLoader resourceLoader) {
        this.resourceLoader = resourceLoader;
        this.properties = new Properties();
    }

    @Override
    public void load(String location) throws Exception {

    }

    @Override
    public void load(String... locations) throws Exception {
        final List<Resource> resources = getResourceLoader().getResource(locations);
        for (Resource resource : resources) {
            InputStream inputStream = resource.getInputStream();
            loadProperties(inputStream);
        }
    }

    /**
     * 加载配置文件
     * @param inputStream 数据流
     * @throws OfficeException
     */
    private void loadProperties(InputStream inputStream) throws OfficeException {
        try {
            getProperties().load(inputStream);
            IOUtils.closeQuietly(inputStream);
        } catch (IOException e) {
            throw new OfficeException("Failed to load property file",e);
        }
    }
}
