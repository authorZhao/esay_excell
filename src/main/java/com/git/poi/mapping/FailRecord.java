package com.git.poi.mapping;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * 失败记录
 */
@Data
@Accessors(chain = true)
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
}
