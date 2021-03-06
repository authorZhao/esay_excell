package com.git.poi.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.util.List;

/**
 * poi工具类
 */
public class PoiUtil {
    /**
     * ；两行复制
     * @param row
     * @param row1
     */
    public static void copyPoiRow(Row row, Row row1) {
        for (int i = 0; i < row.getLastCellNum() ; i++) {
            Cell cell0 = row.getCell(i);
            if(cell0==null)cell0=row.createCell(i);
            Cell cell1 = row1.createCell(i);
            CellType cellType = cell0.getCellType();
            switch (cellType){
                case NUMERIC:cell1.setCellValue(cell0.getNumericCellValue());break;
                case STRING:cell1.setCellValue(cell0.getStringCellValue());break;
                case FORMULA:cell1.setCellValue(cell0.getCellFormula());break;
                case BOOLEAN:cell1.setCellValue(cell0.getBooleanCellValue());break;
                default: cell1.setCellValue("");break;
            }
        }
    }

    /**
     * 数据有效性验证变成字符串输出
     * @param str
     * @return
     */
    public static String getString(String[] str){
        String ss = "";
        for (String s:str) {
            if(ss.equals("")){
                ss += s;
            }else{
                ss += "、"+s;
            }
        }
        return ss;
    }
    public static String getString(List<String> list){
        return getString((String[])list.toArray());
    }


}
