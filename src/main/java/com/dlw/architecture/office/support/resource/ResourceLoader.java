package com.dlw.architecture.office.support.resource;


import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.Properties;

/**
 * @author dengliwen
 * @date 2020/7/6
 * @desc 资源加载器
 * @since 4.0.0
 */
public class ResourceLoader {

    /**
     * 获取参数指定位置的资源 使用场景：获取classpath下约定的文件名内容
     * 类似于 springboot META-INF/spring.factories机制
     * @param location 资源位置
     * @return
     * @throws IOException
     */
    public List<Resource> getResource(String location) throws IOException {
        final Enumeration<URL> urls = this.getClass().getClassLoader().getResources(location);
        List<Resource> resources = new ArrayList<>();
        while(urls.hasMoreElements()) {
            URL url = urls.nextElement();
            final UrlResource resource = new UrlResource(url);
            resources.add(resource);
        }
        return resources;
    }

    /**
     * 获取多个配置文件内容
     * @param locations 多个配置文件
     * @return
     */
    public List<Resource> getResource(String[] locations) throws IOException {
        List<Resource> resources = new ArrayList<>();
        for (String location : locations) {
            resources.addAll(getResource(location));
        }
        return resources;
    }

    /**
     * 获取单个配置文件内容
     * @param location 文件路径
     * @return 文件内容
     */
    public Resource getSingleResource(String location) {
        URL url = this.getClass().getClassLoader().getResource(location);
        return new UrlResource(url);
    }
}
