package com.git.poi.validator;

/**
 * java类型和excell类型映射
 * TODO
 */
public enum FieldType {
    NUMERIC(0,"数值型"),
    STRING(1,"字符串型"),
    FORMULA(2,"公式型"),
    BLANK(3,"空值"),
    BOOLEAN(4,"布尔型"),
    ERROR(5,"错误");

    FieldType(Integer value,String desc) {
        this.value=value;
        this.desc=desc;
    }
    private Integer value;
    private String desc;
}
