package ru.drvsh.rebus.bean;

public class SpecificationBean {
    private final String specification;
    private final String requirement;
    private final String idClause;
    private final String clause;

    public SpecificationBean(String specification, String requirement, String idClause, String clause) {
        this.specification = specification;
        this.requirement = requirement;
        this.idClause = idClause;
        this.clause = clause;
    }
}
