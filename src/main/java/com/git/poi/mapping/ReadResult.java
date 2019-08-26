package com.git.poi.mapping;

import lombok.Data;
import lombok.experimental.Accessors;

import java.util.List;

/**
 * 读取的结果对象
 * @param <T>
 */
@Data
@Accessors(chain = true)
public class ReadResult<T> {
    /**
     * 读取成功的集合
     */
    List<T> sucesslist;

    /**
     * 失败集合
     */
    List<FailRecord> failRecordList;

    /**
     * 总记录数
     */
    Integer total;

}
