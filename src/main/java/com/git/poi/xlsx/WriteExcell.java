package com.git.poi.xlsx;

import com.git.poi.mapping.ExcelMapping;
import com.git.poi.mapping.FailRecord;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

public interface WriteExcell {

    Integer MAX_ROW = 50000;

    /**
     * 根据失败记录返回
     * @param file xlsx文件
     * @param failRecordList 失败记录
     * @return
     */
    Workbook failRecordsReturn(File file, List<FailRecord> failRecordList) throws Exception;

    /**
     * 根据集合生成workBook
     * @param list 要导出的集合
     * @param excelMapping
     * @return
     */
    Workbook generateXlsxWorkbook(List list,ExcelMapping excelMapping);

    Workbook generateXlsxWorkbook(List list,ExcelMapping excelMapping,Integer maxRow);
}
