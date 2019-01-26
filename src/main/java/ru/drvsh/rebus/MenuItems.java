package ru.drvsh.rebus;

public class MenuItems {
    /** №пп */
    private String idItemName;
    /** Наименование товара */
    private String nameItemName;
    /** Наименование показателя (характеристики) товара */
    private String specificationItemName;
    /** Требование к показателю (характеристике) товара */
    private String requirementItemName;
    /** Вид и подвид показателя (характеристики) товара */
    private String idClauseItemName;

    public MenuItems(String idItemName, String nameItemName, String specificationItemName, String requirementItemName, String idClauseItemName) {
        this.idItemName = idItemName;
        this.nameItemName = nameItemName;
        this.specificationItemName = specificationItemName;
        this.requirementItemName = requirementItemName;
        this.idClauseItemName = idClauseItemName;
    }

    public String getIdItemName() {
        return idItemName;
    }

    public String getNameItemName() {
        return nameItemName;
    }

    public String getSpecificationItemName() {
        return specificationItemName;
    }

    public String getRequirementItemName() {
        return requirementItemName;
    }

    public String getIdClauseItemName() {
        return idClauseItemName;
    }
}
