package com.git.poi.util;

import com.git.poi.exception.CheckException;
import com.git.poi.mapping.ExcelProperty;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;

/**
 * 读取工具类
 */
public class ReadUtil {

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

    private static <T> void cellValueFormat(T t, XSSFCell cell, Method m2, Field field, String checkMethod) {
    }


}
