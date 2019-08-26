package com.git.poi.mapping;

import lombok.Data;
import lombok.experimental.Accessors;

import java.util.List;

/**
 * 用于记录某个对象对应的行数
 */
@Data
@Accessors(chain = true)
public class BaseRow {

    List<Integer> rowNums;
}
