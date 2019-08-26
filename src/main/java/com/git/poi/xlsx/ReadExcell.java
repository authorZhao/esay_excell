package com.git.poi.xlsx;

import com.git.poi.mapping.ReadResult;

import java.io.File;
import java.util.List;
import java.util.Map;

public interface ReadExcell<T> {
    /**
     * 读取cexll数据
     * @param f
     * @param checkMap
     * @return
     */
    ReadResult<T> readExcell(File f, Map<Integer,List> checkMap);

    /**
     *
     * @param f xlsx文件
     * @return
     */
    ReadResult<T> readExcell(File f);
}
