package com.dlw.architecture.office.model.pdf;

import lombok.*;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * @author dengliwen
 * @date 2020/6/18
 * @desc PDF table对象数据封装
 * @since 4.0.0
 */
@NoArgsConstructor
public class PdfTableModel {

    /**
     * 用于多个table的展示顺序
     */
    @Getter
    @Setter
    private int index;

    /**
     * 导出PDF模型class
     */
    @Getter
    @Setter
    private Class<?> clazz;

    /**
     * 一个table的数据
     */
    @Getter
    @Setter
    private Collection<?> data;

    private transient List<PdfTableModel> models = new ArrayList<>();

    public PdfTableModel(int index, Class<?> clazz, Collection<?> data) {
        this.index = index;
        this.data = data;
        this.clazz = clazz;
    }

    public static PdfTableModel builder() {
        return new PdfTableModel();
    }

    public List<PdfTableModel> list() {
        return this.models;
    }

    public PdfTableModel of(int index,Class<?> clazz,Collection<?> data) {
        final PdfTableModel model = new PdfTableModel(index,clazz, data);
        this.models.add(model);
        return this;
    }


}
