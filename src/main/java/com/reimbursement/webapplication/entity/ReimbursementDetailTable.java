package com.reimbursement.webapplication.entity;

import java.util.ArrayList;
import java.util.List;

/*
* 費用明細表
* */
public class ReimbursementDetailTable {

    private String startDate; //開始時間

    private String endDate; //結束時間

    private String travelDate; //出差時間

    private String travelReason;//出差理由

    private List<CostDetail> costTypeList = new ArrayList<CostDetail>(); //費用類型

    private  String  advance; //是否提前出發

    private String projectName; //項目名稱

    public String getStartDate() {
        return startDate;
    }

    public String getEndDate() {
        return endDate;
    }

    public String getTravelDate() {
        return travelDate;
    }

    public String getTravelReason() {
        return travelReason;
    }

    public void setStartDate(String startDate) {
        this.startDate = startDate;
    }

    public void setEndDate(String endDate) {
        this.endDate = endDate;
    }

    public void setTravelDate(String travelDate) {
        this.travelDate = travelDate;
    }

    public void setTravelReason(String travelReason) {
        this.travelReason = travelReason;
    }

    public String getProjectName() {
        return projectName;
    }

    public void setProjectName(String projectName) {
        this.projectName = projectName;
    }

    public List<CostDetail> getCostTypeList() {
        return costTypeList;
    }

    public void setCostTypeList(List<CostDetail> costTypeList) {
        this.costTypeList = costTypeList;
    }

    public String getAdvance() {
        return advance;
    }

    public void setAdvance(String advance) {
        this.advance = advance;
    }
}