package com.git.poi.mapping;

/**
 * 失败记录
 */
public class FailRecord {
    /**
     * 失败的行数，下标
     */
    private Integer rowNum;
    /**
     * 失败或者异常信息
     */
    private String message;
    /**
     * 失败的列数，下标
     */
    private Integer colNum;

    public Integer getRowNum() {
        return rowNum;
    }

    public FailRecord setRowNum(Integer rowNum) {
        this.rowNum = rowNum;
        return this;
    }

    public String getMessage() {
        return message;
    }

    public FailRecord setMessage(String message) {
        this.message = message;
        return this;
    }

    public Integer getColNum() {
        return colNum;
    }

    public FailRecord setColNum(Integer colNum) {
        this.colNum = colNum;
        return this;
    }
}
