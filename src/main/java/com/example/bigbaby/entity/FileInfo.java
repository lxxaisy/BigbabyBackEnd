package com.example.bigbaby.entity;

/**
 * 用户登录对象
 * 
 * @author easycloud
 */
public class FileInfo
{
    /**
     * orderNo
     */
    private String orderNo;

    /**
     * poNo
     */
    private String poNo;

    /**
     * count
     */
    private double count;

    /**
     * enName
     */
    private String enName;

    /**
     * contractName
     */
    private String contractName;

    /**
     * fileName
     */
    private String fileName;

    /**
     * type
     */
    private String type;

    public String getOrderNo() {
        return orderNo;
    }

    public void setOrderNo(String orderNo) {
        this.orderNo = orderNo;
    }

    public String getPoNo() {
        return poNo;
    }

    public void setPoNo(String poNo) {
        this.poNo = poNo;
    }

    public double getCount() {
        return count;
    }

    public void setCount(double count) {
        this.count = count;
    }

    public String getEnName() {
        return enName;
    }

    public void setEnName(String enName) {
        this.enName = enName;
    }

    public String getContractName() {
        return contractName;
    }

    public void setContractName(String contractName) {
        this.contractName = contractName;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }
}
