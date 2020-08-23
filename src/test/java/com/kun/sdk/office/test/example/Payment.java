package com.kun.sdk.office.test.example;

import com.dlw.architecture.office.annotation.WordLoopTable;
import com.dlw.architecture.office.annotation.WordMiniTable;
import com.dlw.architecture.office.annotation.WordPicture;
import com.dlw.architecture.office.annotation.WordText;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

/**
 * @author dengliwen
 * @date 2020/7/13
 * @desc
 */
@Data
public class Payment {

    @WordText
    private String no = "N-0001";

    @WordText
    private String id = "001";
    @WordPicture(width = 30,height = 30)
    private String img = "http://deepoove.com/images/icecream.png";
    @WordText
    private String title = "xxx公司";
    @WordText
    private String receive = "甲方";

    @WordMiniTable(header = {"日期","订单编号","销售代表","价格","发货方式","税号"})
    private List<Order> orders = new ArrayList<>();

    @WordLoopTable
    private List<Detail> details = new ArrayList<>();

    @WordLoopTable
    private List<Labor> labors = new ArrayList<>();

    @WordText
    private double subTotal = 10000;

    @WordText
    private double tax = 1000;

    @WordText
    private double transform = 5000;

    @WordText
    private double other = 4000;

    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    private static class Order {
        private String date;
        private String orderNo;
        private String sale;
        private String price;
        private String way;
        private String taxID;
    }

    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    private static class Detail {
        private int count;
        private String goods;
        private String remark;
        private double discount;
        private double tax;
        private double price;
        private double total;

    }

    @Data
    @AllArgsConstructor
    private static class Labor {
        private String category;
        private int people;
        private int price;
        private int total;

        public Labor(String category, int people, int price) {
            this.category = category;
            this.people = people;
            this.price = price;
        }
    }

    private void init() {
        for (int i = 0; i < 3; i++) {
            Order order = new Order("2020-07-01",UUID.randomUUID().toString(),"张三","200元","快递","T000" + i);
            orders.add(order);
        }

        for (int i = 1; i < 5; i++) {
            Detail detail = new Detail(i * 2,"冰淇淋","",0.4,0.5,10,i * 2 * 10);
            details.add(detail);
        }

        for (int i = 0; i < 3; i++) {
            Labor labor = new Labor("种类" + i,i + 1,100,(i+1)*100);
            labors.add(labor);
        }
    }

    public Payment() {
        init();
    }
}
