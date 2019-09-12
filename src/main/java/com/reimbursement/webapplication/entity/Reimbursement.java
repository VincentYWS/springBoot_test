package com.reimbursement.webapplication.entity;

/**
 * 費用
 * */
public class Reimbursement {

    private String Date; //時間

    private String Description; //描述信息

    private String Currency; //貨幣類型

    private double Amount ; //金額

    private String Allowance; //excel表類型

    private String projectName; //項目名稱

    private String project; //項目


    public String getDate() {
        return Date;
    }

    public String getDescription() {
        return Description;
    }

    public String getCurrency() {
        return Currency;
    }

    public String getAllowance() {
        return Allowance;
    }

    public void setDate(String date) {
        Date = date;
    }

    public void setDescription(String description) {
        Description = description;
    }

    public void setCurrency(String currency) {
        Currency = currency;
    }

    public void setAllowance(String allowance) {
        Allowance = allowance;
    }

    public String getProjectName() {
        return projectName;
    }

    public void setProjectName(String projectName) {
        this.projectName = projectName;
    }

    public String getProject() {
        return project;
    }

    public void setProject(String project) {
        this.project = project;
    }

    public double getAmount() {
        return Amount;
    }

    public void setAmount(double amount) {
        Amount = amount;
    }
}