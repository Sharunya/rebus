package ru.drvsh.rebus.bean;

public class ClauseBean {
    private String id;
    private String text;

    public ClauseBean(String id, String text) {
        this.id = id;
        this.text = text;
    }

    public String getId() {
        return id;
    }

    public String getText() {
        return text;
    }

    @Override
    public String toString() {
        return "ClauseBean{" +
                "id='" + id + '\'' +
                ", text='" + text + '\'' +
                '}';
    }
}
