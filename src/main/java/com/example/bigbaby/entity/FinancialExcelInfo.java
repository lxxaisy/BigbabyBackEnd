package com.example.bigbaby.entity;

import lombok.Data;

@Data
public class FinancialExcelInfo {
    /**
     * 汇率
     */
    private String exchangeRate;

    /**
     * 伝票No.
     */
    private String index;

    /**
     * 摘要
     */
    private String abstractInfo;

    /**
     * 借方勘定科目
     */
    private String debitCodeJP;

    /**
     * 借方科目参考
     */
    private String debitCodeName;

    /**
     * 借方科目
     */
    private String debitCode;

    /**
     * 借方補助科目
     */
    private String debitMinorNameJP;

    /**
     * 借方部門
     */
    private String debitMinorDepartmentNameJP;

    /**
     * 借方部门
     */
    private String debitMinorDepartmentName;

    /**
     * 借方税区分
     */
    private String debitMinorTaxInfo;

    /**
     * 借方含税金额
     */
    private String debitCountWithTax;

    /**
     * 借方消費税額
     */
    private String consumptionTax;

    /**
     * 借方去税金额
     */
    private String debitCount;

    /**
     * 贷方勘定科目
     */
    private String creditCodeJP;

    /**
     * 贷方科目参考
     */
    private String creditCodeName;

    /**
     * 贷方科目
     */
    private String creditCode;

    /**
     * 贷方補助科目
     */
    private String creditMinorNameJP;

    /**
     * 贷方部門
     */
    private String creditMinorDepartmentNameJP;

    /**
     * 贷方部门
     */
    private String creditMinorDepartmentName;

    /**
     * 贷方税区分
     */
    private String creditMinorTaxInfo;

    /**
     * 贷方含税金额
     */
    private String creditCountWithTax;

    /**
     * 贷方消費税額
     */
    private String creditConsumptionTax;

    /**
     * 贷方去税金额
     */
    private String creditCount;

    /**
     * 是否为工资 "0": 否 "1": 是
     */
    private String isSalary;
}
