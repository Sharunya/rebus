package ru.drvsh.rebus.bean;

import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

public class ProductBeen {
    private final String id;
    private final String name;
    private final List<SpecificationBean> specificationList;
    private final CellRangeAddress rangeAddress;

    public ProductBeen(CellRangeAddress rangeAddress, String id, String name) {
        this.id = id;
        this.name = name;
        this.specificationList = new ArrayList<>();
        this.rangeAddress = rangeAddress;
    }

    public boolean add(SpecificationBean specificationBean) {
        return specificationList.add(specificationBean);
    }

    public CellRangeAddress getRangeAddress() {
        return rangeAddress;
    }

    public String getId() {
        return id;
    }

    public String getName() {
        return name;
    }

    public List<SpecificationBean> getSpecificationList() {
        return specificationList;
    }

    @Override
    public String toString() {
        return "ProductBeen{" +
                "id='" + id + '\'' +
                ", name='" + name + '\'' +
                ", specificationList=" + specificationList +
                '}';
    }
}
