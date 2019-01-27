package ru.drvsh.rebus.bean;

import java.text.ParseException;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.ibm.icu.text.RuleBasedNumberFormat;

public class SpecificationBean {
    private static final Pattern PATTERN = Pattern.compile("\\d+[.,]*\\d*");

    private static final RuleBasedNumberFormat NUMBER_FORMAT = new RuleBasedNumberFormat(Locale.forLanguageTag("ru"), RuleBasedNumberFormat.SPELLOUT);

    private final String specificationRaw;
    private final String specification;
    private final String requirementRaw;
    private final String requirement;
    private final String idClause;
    private final String clauseRaw;
    private final String clause;

    public SpecificationBean(String specification, String requirement, String idClause, String clause) throws ParseException {
        this.specificationRaw = specification;
        this.specification = format(specification);
        this.requirementRaw = requirement;
        this.requirement = format(requirement);
        this.idClause = idClause;
        this.clauseRaw = clause;
        this.clause = format(clause);
    }

    public String getSpecification() {
        return specification;
    }

    public String getRequirement() {
        return requirement;
    }

    public String getIdClause() {
        return idClause;
    }

    public String getClause() {
        return clause;
    }

    @Override
    public String toString() {
        return "SpecificationBean{" + "specification='" + specification + '\'' + ", requirement='" + requirement + '\'' + ", idClause='" + idClause + '\'' + ", clause='" + clause + '\'' + '}';
    }

    private String format(String input) throws ParseException {
        if(true) return input;
        Matcher matcher = PATTERN.matcher(input + " ");
        StringBuilder result = new StringBuilder();
        int start = 0;
        while (matcher.find()) {
            int first = matcher.start();
            int last = matcher.end();

            result.append(input, start, first);
            result.append(matcher.group()).append(" (");
            result.append(NUMBER_FORMAT.format(NUMBER_FORMAT.parse(matcher.group()))).append(") ");
            start = last;
        }
        result.append(input.substring(start));

        return result.toString();
    }
}
