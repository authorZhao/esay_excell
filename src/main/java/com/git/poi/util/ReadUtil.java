package com.git.poi.util;

import com.git.poi.exception.CheckException;
import com.git.poi.mapping.ExcelProperty;
import com.git.poi.validator.Check;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.List;

/**
 * 读取工具类
 */
public class ReadUtil {

    /**
     * 将某一行转化为对应的对象
     * @param clazz 该对象的字节码对象
     * @param row 行
     * @param propertyList 属性list
     * @param <T> 该对象
     * @return <T>
     * @throws Exception
     */
    public static <T> T mappingRowToList(Class<T> clazz,XSSFRow row, List<ExcelProperty> propertyList) throws Exception{
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
     *
     * @param t 返回的对象
     * @param cell 列
     * @param m2 该对象的set方法
     * @param field 需要set的字段
     * @param checkMethod 校验的方法名称
     * @param <T>
     */
    private static <T> void cellValueFormat(T t, XSSFCell cell, Method m2, Field field, String checkMethod) throws CheckException, InvocationTargetException, IllegalAccessException {
        //TODO
        if(cell==null)return;
        switch (cell.getCellType()){
            case NUMERIC:
                Number num = cell.getNumericCellValue();
                if(StringUtils.isNotBlank(checkMethod)){
                    String s = num.toString();
                    Boolean flag=false;
                    try {
                        flag = check(s,checkMethod);
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
                if(StringUtils.isNotBlank(checkMethod)){
                    Boolean aaa=false;
                    try {
                        aaa = check(value,checkMethod);
                    } catch (Exception e) {
                        throw new CheckException("字段校验异常");
                    }
                    if(!aaa)throw new CheckException("字段校验异常");
                }
                break;
            case BOOLEAN:
                m2.invoke(t,cell.getBooleanCellValue());
                break;
            default:
        }
    }

    /**
     *
     * @param m2 setxxx方法
     * @param t  该对象
     * @param num 数字类型
     * @param field 该对象该字段的类型
     * @param <T> 返回void
     */
    private static <T> void numConvet(Method m2, T t, Number num, Field field) throws InvocationTargetException, IllegalAccessException {
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


    /**
     *
     * @param s 需要校验的字符串
     * @param checkMethodName  checkMethod方法的名称
     * @return
     */
    private static Boolean check(String s, String checkMethodName) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        Method checkMethod = Check.class.getMethod(checkMethodName,String.class);
        boolean flag = (Boolean)checkMethod.invoke(null,s);
        return flag;
    }


}
