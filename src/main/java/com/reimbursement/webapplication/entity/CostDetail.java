package com.reimbursement.webapplication.entity;

/*
* 費用類型
* */
public class CostDetail {

    private String costType ; //費用類型

    private String currency; //錢類型

    private String money ; // 多少錢

    public String getCurrency() {
        return currency;
    }

    public void setCurrency(String currency) {
        this.currency = currency;
    }

    public String getMoney() {
        return money;
    }

    public void setMoney(String money) {
        this.money = money;
    }

    public String getCostType() {
        return costType;
    }

    public void setCostType(String costType) {
        this.costType = costType;
    }
}