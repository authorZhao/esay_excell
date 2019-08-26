package com.git.poi.xlsx;

import com.git.poi.exception.CheckException;
import com.git.poi.exception.FileException;
import com.git.poi.mapping.ExcelMapping;
import com.git.poi.mapping.ExcelProperty;
import com.git.poi.mapping.FailRecord;
import com.git.poi.mapping.MappingFactory;
import com.git.poi.validator.Check;
import com.git.poi.validator.Options;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * 读取数据
 * @author authorZhao
 */

public class ReadXlsxExcell<T> {
    private static final Logger logger = LoggerFactory.getLogger(ReadXlsxExcell.class);

    private final Class<?> clazz;
    private final ExcelMapping excelMapping;

    private List<ExcelProperty> propertyList;

    public ReadXlsxExcell(Class<T> clazz){
        this.clazz = clazz;
        this.excelMapping = MappingFactory.get(this.clazz);
    }

    /**
     * 读取只读第一张表
     * @param f xlsx文件
     * @param checkMap 校验的list，以序号命名
     * @return data:成功数据
     *         totalNum 总行数
     *         faileList FailRecord
     *         reason：失败原因
     *
     * @throws Exception
     */
    public Map<String,Object> getData(File f,Map<Integer,List> checkMap) throws Exception {
        List<FailRecord> faileList = new ArrayList<>();
        Map<String,Object> map = new HashMap<>();
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
                t = mappingRowToList(sheet.getRow(i), excelMapping.getPropertyList());
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
        map.put("data",list);
        map.put("totalNum",sheet.getLastRowNum());
        map.put("faileList",faileList);
        return map;
    }

    public Map<String,Object> getData(File f) throws Exception {
        return this.getData(f,null);
    }

    /**
     * 转换
     * @param t
     * @param excelProperty
     * @param checkMap
     * @throws Exception
     */
    private void exchage(T t, ExcelProperty excelProperty, Map<Integer, List> checkMap) throws Exception {
        if(checkMap==null)return;
        Integer key = excelProperty.getSort();
        List checkList = checkMap.get(key);
        Options.getString(checkList,excelProperty,t);
    }

    /**
     * 校验
     * @param t
     * @param checkMethodName
     * @throws Exception
     */
    private Boolean check(String t, String checkMethodName) throws Exception {
        Method checkMethod = Check.class.getMethod(checkMethodName,String.class);
        boolean flag = (Boolean)checkMethod.invoke(null,t);
        return flag;
    }

    /**
     *
     * @param row xlsx文件的row，包含每条记录
     * @param propertyList 导入的属性配置
     * @return T
     * @throws Exception
     */
    private T mappingRowToList(XSSFRow row, List<ExcelProperty> propertyList) throws Exception {
        T t = null;
        try {
            t = (T)clazz.newInstance();
        } catch (ReflectiveOperationException e) {
            throw new ReflectiveOperationException(clazz.getName()+"对象不能实例化");
        }
        for (int i = 0; i < propertyList.size(); i++) {

            XSSFCell cell = row.getCell(i);
            String methodName = "set" + (new StringBuilder()).append(Character.toUpperCase(propertyList.get(i).getColumn().charAt(0))).append(propertyList.get(i).getColumn().substring(1)).toString();
            Field field = null;
            Method m2 = null;
            String column = propertyList.get(i).getColumn();
            //boolean require = propertyList.get(i).getRequire();
            try {
                field = clazz.getDeclaredField(column);
                m2 = clazz.getDeclaredMethod(methodName, field.getType());
                cellValueFormat(t, cell, m2, field,propertyList.get(i).getCheckMethod());

            } catch (NoSuchFieldException e) {
                throw new NoSuchFieldException(propertyList.get(i).getColumn()+"字段不存在_qdz"+i);
            } catch (NoSuchMethodException e) {
                throw new NoSuchMethodException(methodName+"方法不存在_qdz"+i);
            } catch (ReflectiveOperationException e) {
                throw new ReflectiveOperationException("参数设置异常_qdz"+i);
            }catch (CheckException e){
                throw new CheckException(propertyList.get(i).getColumn()+"字段校验异常_qdz"+i);
            };
        }
        return t;

    }

    /**
     * 值得映射
     * @param t
     * @param cell
     * @param m2
     * @param field
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     */
    private void cellValueFormat (T t, XSSFCell cell, Method m2, Field field, String chechMethod) throws InvocationTargetException, IllegalAccessException,CheckException{
        //TODO
        if(cell==null)return;
        switch (cell.getCellType()){
            case NUMERIC:
                Number num = cell.getNumericCellValue();
                if(StringUtils.isNotBlank(chechMethod)){
                    cell.setCellType(CellType.STRING);
                    cell.setCellStyle(new XSSFCellStyle(new StylesTable()));
                    String s = cell.getStringCellValue();
                    Boolean flag=false;
                    try {
                        flag = check(s,chechMethod);
                    } catch (Exception e) {
                        throw new CheckException("字段校验异常");
                    }
                    if(!flag)throw new CheckException("字段校验异常");
                }
                if(num instanceof Long) {
                    numConvet(m2,t,num,field);
                    break;
                }
                if(num instanceof Integer) {
                    numConvet(m2,t,num,field);
                    break;
                }
                if(num instanceof Double) {
                    numConvet(m2,t,num,field);
                    break;
                }
                if(num instanceof BigDecimal) {
                    numConvet(m2,t,num,field);
                    break;
                }
                break;
            case STRING:
                String value = cell.getStringCellValue();
                m2.invoke(t,value);
                if(StringUtils.isNotBlank(chechMethod)){
                    Boolean aaa=false;
                    try {
                        aaa = check(value,chechMethod);
                    } catch (Exception e) {
                        throw new CheckException("字段校验异常");
                    }
                    if(!aaa)throw new CheckException("字段校验异常");
                }
                break;
            case BOOLEAN:
                m2.invoke(t,cell.getBooleanCellValue());
                break;
            case BLANK:
                break;

            case _NONE:
                break;
            case FORMULA:
                break;
            case ERROR:
                break;
            default:
                break;
        }

    }

    /**
     * 数字转换
     * @param m2
     * @param t
     * @param num
     * @param field
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     */
    private void numConvet(Method m2, T t, Number num, Field field) throws InvocationTargetException, IllegalAccessException {

        if(field.getType()== Long.class){
            m2.invoke(t,num.longValue());
        }
        if(field.getType()== Integer.class){
            m2.invoke(t,num.intValue());
        }
        if(field.getType()== Double.class){
            m2.invoke(t,num.doubleValue());
        }
        if(field.getType()== BigDecimal.class){
            m2.invoke(t,BigDecimal.valueOf(num.doubleValue()));
        }
        if(field.getType()== String.class){
            BigDecimal bd = new BigDecimal((Double)num);
            String str = bd.toPlainString();
            m2.invoke(t,str);
        }


    }


}

