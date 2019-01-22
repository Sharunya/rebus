package ru.drvsh.rebus.bean;

import java.util.List;

public class ProductBeen {
    private String id;
    private String name;
    private List<SpecificationBean> specificationList;

    public ProductBeen(String id, String name) {
        this.id = id;
        this.name = name;
    }
}
