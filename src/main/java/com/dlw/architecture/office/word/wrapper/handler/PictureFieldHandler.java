package com.dlw.architecture.office.word.wrapper.handler;

import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.util.BytePictureUtils;
import com.dlw.architecture.office.annotation.WordPicture;
import com.dlw.architecture.office.exception.OfficeException;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.io.InputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.Map;

/**
 * @author dengliwen
 * @date 2020/6/12
 * @desc 图片字段类型处理器
 * @since 4.0.0
 */
public class PictureFieldHandler<T> extends AbstractFieldHandler<T> {


    public PictureFieldHandler(Field field) {
        super(field);
    }

    @Override
    public void handle(Map<String, Object> result, Field field, T t, Annotation annotation) throws OfficeException {
        WordPicture wordPicture = (WordPicture) annotation;
        String fieldName = field.getName();
        final Object value = getFieldValue(field, t);
        if (StringUtils.isNotBlank(this.alias)) {
            fieldName = this.alias;
        }
        //表示图片本地或网络地址
        if (value instanceof String) {
            String path = (String) value;
            //网络图片
            if (path.startsWith("http")) {
                result.put(fieldName,new PictureRenderData(wordPicture.width(),wordPicture.height(),
                        wordPicture.format(), BytePictureUtils.getUrlBufferedImage(path)));
            }else {
                //本地图片
                result.put(fieldName,new PictureRenderData(wordPicture.width(),wordPicture.height(),path));
            }
        }else if (value instanceof File) {
            File file = (File) value;
            result.put(fieldName,new PictureRenderData(wordPicture.width(),wordPicture.height(),file));
        }else if (value instanceof InputStream) {
            InputStream inputStream = (InputStream) value;
            result.put(fieldName,new PictureRenderData(wordPicture.width(),wordPicture.height(),
                    wordPicture.format(),inputStream));
        }
    }
}
