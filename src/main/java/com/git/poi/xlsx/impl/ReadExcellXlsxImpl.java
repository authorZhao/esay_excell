package com.git.poi.xlsx.impl;

import com.git.poi.exception.FileException;
import com.git.poi.mapping.ExcelMapping;
import com.git.poi.mapping.FailRecord;
import com.git.poi.mapping.MappingFactory;
import com.git.poi.mapping.ReadResult;
import com.git.poi.util.ReadUtil;
import com.git.poi.xlsx.ReadExcell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 读取xlsx的实现类
 * @param <T>
 */
public class ReadExcellXlsxImpl<T> implements ReadExcell {

    private final Class<?> clazz;
    private final ExcelMapping excelMapping;

    public ReadExcellXlsxImpl(Class<T> clazz){
        this.clazz = clazz;
        this.excelMapping = MappingFactory.get(this.clazz);
    }

    @Override
    public ReadResult<T> readExcell(File f, Map checkMap){
        ReadResult<T> readResult = new ReadResult<>();
        List<FailRecord> faileList = new ArrayList<>();
        List<T> list = new ArrayList<>();
        XSSFWorkbook workbook = null;
        try {
            workbook = new XSSFWorkbook(new FileInputStream(f));
        } catch (IOException e) {
            e.printStackTrace();
            throw new FileException("无法读取该xlsx文件");
        }
        XSSFSheet sheet = workbook.getSheetAt(0);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {//从第二行开始读取
            T t = null;
            try {
                t = ReadUtil.mappingRowToList((Class<T>) clazz,sheet.getRow(i), excelMapping.getPropertyList());
                Method m2 = clazz.getDeclaredMethod("setRowNums", List.class);
                List<Integer> rowList = new ArrayList<>();
                rowList.add(i);
                m2.invoke(t,rowList);
                //exchage(t,excelMapping.getPropertyList().get(0),checkMap);
            }catch (Exception e){
                FailRecord failRecord = new FailRecord();
                String message = e.getMessage();
                if(message==null)message="";
                String[] msg = message.split("_qdz",2);
                if(msg.length>=2){
                    failRecord.setRowNum(i).setMessage(msg[0]).setColNum(Integer.valueOf(msg[1]));
                }else{
                    failRecord.setRowNum(i).setMessage(msg[0]);
                }
                faileList.add(failRecord);
            }
            if(t!=null)list.add(t);
        }
        return readResult.setSucesslist(list).setFailRecordList(faileList).setTotal(sheet.getLastRowNum());

    }


    @Override
    public ReadResult<T> readExcell(File f) {
        return readExcell(f,null);
    }
}
