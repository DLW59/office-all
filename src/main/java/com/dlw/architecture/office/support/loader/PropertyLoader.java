package com.dlw.architecture.office.support.loader;

/**
 * @author dengliwen
 * @date 2020/7/6
 * @desc 属性解析器
 * @since 4.0.0
 */
public interface PropertyLoader {

    /**
     * 加载配置文件
     * 约定配置文件为 eg: META-INF/xxx.xxx
     * @param location 文件约定的路径
     * @throws Exception
     */
    void load(String location) throws Exception;

    /**
     * 加载配置文件
     * @param locations 多个配置文件路径
     * @throws Exception
     */
    void load(String... locations) throws Exception;
}
