package com.reli.PhishingReportGenerator;

public class Employee {
    private String firstname;
    private String lastname;
    private String email;
    private boolean open;
    private boolean click;
    private int multiClick;
    private boolean reported;
    private boolean bounced;
    private boolean pass;
    private String template;

    public Employee(String firstname, String lastname, String email, boolean open, boolean click, int multiClick, boolean reported, boolean bounced, boolean pass, String template) {
        this.firstname = firstname;
        this.lastname = lastname;
        this.email = email;
        this.open = open;
        this.click = click;
        this.multiClick = multiClick;
        this.reported = reported;
        this.bounced = bounced;
        this.pass = pass;
        this.template = template;
    }

    public String getFirstname() {
        return firstname;
    }

    public String getLastname() {
        return lastname;
    }

    public String getEmail() {
        return email;
    }

    public boolean isOpen() {
        return open;
    }

    public boolean isClick() {
        return click;
    }

    public int getMultiClick() {
        return multiClick;
    }

    public boolean isReported() {
        return reported;
    }

    public boolean isBounced() {
        return bounced;
    }

    public boolean isPass() {
        return pass;
    }

    public String getTemplate() {
        return template;
    }
}
