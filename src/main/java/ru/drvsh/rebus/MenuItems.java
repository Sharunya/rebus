package ru.drvsh.rebus;

public class MenuItems {
    private String idItemName;
    private String nameItemName;
    private String specificationItemName;
    private String requirementItemName;
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
