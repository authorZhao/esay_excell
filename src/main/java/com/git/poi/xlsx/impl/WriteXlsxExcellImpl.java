package com.git.poi.xlsx.impl;

import com.git.poi.mapping.ExcelMapping;
import com.git.poi.mapping.ExcelProperty;
import com.git.poi.mapping.FailRecord;
import com.git.poi.util.WriteUtil;
import com.git.poi.xlsx.WriteExcell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;

import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;

public class WriteXlsxExcellImpl implements WriteExcell {
    private static final Logger logger = LoggerFactory.getLogger(WriteXlsxExcellImpl.class);

    @Override
    public Workbook failRecordsReturn(File file, List<FailRecord> failRecordList) throws IOException {
        if(logger.isInfoEnabled())logger.info("开始返回失败数据");
        XSSFWorkbook workbook  = new XSSFWorkbook(new FileInputStream(file));
        if(failRecordList==null||failRecordList.size()<=0)return workbook;
        //只拿第一张,或者循环
        XSSFSheet sheet = workbook.getSheetAt(0);
        List<Integer> failRows = failRecordList.stream().map(FailRecord::getRowNum).sorted().collect(Collectors.toList());
        failRows.add(0);//第一行保留
        List<Integer> totalRows = new ArrayList<>(sheet.getLastRowNum());
        for (int i = 0; i < sheet.getLastRowNum() ; i++) {
            totalRows.add(i);
        }

        //为失败记录添加原因
        for (int i = 0; i < failRecordList.size(); i++) {
            FailRecord failRecord = failRecordList.get(i);
            Row row = sheet.getRow(failRecord.getRowNum());
            Cell cell = row.createCell(row.getLastCellNum()+1);
            cell.setCellValue(failRecord.getMessage());
        }
        totalRows.removeAll(failRows);
        totalRows.stream().forEach(t->sheet.removeRow(sheet.getRow(t)));
        if(logger.isInfoEnabled())logger.info("正在返回失败数据");
        return workbook;
    }

    @Override
    public Workbook generateXlsxWorkbook(List list, ExcelMapping excelMapping) {
        return generateXlsxWorkbook(list,excelMapping,MAX_ROW);
    }



    @Override
    public Workbook generateXlsxWorkbook(List list, ExcelMapping excelMapping, Integer maxRow) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        //数据简单校验
        if(excelMapping==null)return workbook;
        if(list==null||list.isEmpty())return workbook;
        List<ExcelProperty> propertyList = excelMapping.getPropertyList();
        Double sheetNo =1.0;
        if(list!=null&&list.size()>0)sheetNo = Math.ceil(list.size() / (double) maxRow);
        for (int index = 0; index <= (sheetNo == 0.0 ? sheetNo : sheetNo - 1); index++) {
            String sheetName = excelMapping.getName() + (index == 0 ? "" : "_" + index);
            //设置表头
            XSSFSheet sheet = WriteUtil.setExcellXlsxHeader(workbook, propertyList,sheetName);
            //数据校验
            //setValidation(mExcelMapping.getPropertyList().get(3),sheet,new String[]{"是","否"},3);
            //设置表身
            int startNo = index * maxRow;
            int endNo = Math.min(startNo + maxRow, list.size());
            //从前两个对象里面取出，可能都不存在
            Class clazz = list.get(0).getClass()==null?list.get(0).getClass():list.get(1).getClass();
            for (int i = startNo; i < endNo; i++) {
                Row bodyRow = sheet.createRow(i-startNo+1);
                for (int j = 0; j < propertyList.size(); j++) {
                    Cell cell = bodyRow.createCell(j);
                    //设置值
                    Object obj = null;
                    String methodName = "get"+(new StringBuilder()).append(Character.toUpperCase(propertyList.get(j).getColumn().charAt(0))).append(propertyList.get(j).getColumn().substring(1)).toString();
                    try{
                        Method m2 = clazz.getDeclaredMethod(methodName);
                        obj = m2.invoke(list.get(i));
                    } catch (Exception e) {
                        logger.error("反射处理失败");
                        e.printStackTrace();
                    }
                    if(obj instanceof String)cell.setCellValue((String)obj);
                    if(obj instanceof Number){
                        if(obj instanceof Long)cell.setCellValue((Long)obj);
                        if(obj instanceof Double)cell.setCellValue((Double)obj);
                        if(obj instanceof Integer)cell.setCellValue((Integer)obj);
                    }
                }
            }
        }
        return workbook;
    }

}
