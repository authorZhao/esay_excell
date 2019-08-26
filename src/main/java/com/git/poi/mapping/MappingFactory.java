package com.git.poi.mapping;

import com.git.poi.annotation.Excel;
import com.git.poi.annotation.ExcelField;
import com.git.poi.exception.AnnotationNotExistException;
import com.git.poi.exception.ClassMappingException;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;

import static java.util.stream.Collectors.toList;

public class MappingFactory {

    /**
     * 获取指定实体的Excel映射信息
     *
     * @param clazz 实体
     * @return ExcelMapping映射对象
     */
    public static ExcelMapping get(Class<?> clazz) {
        try {
            //return MappingFactory.getFromXmlOrAnno(clazz);
            return getFromXmlOrAnno(clazz);
        } catch (Exception e) {
            throw new ClassMappingException("class对象不存在");
        }
    }

    /**
     * 读取注解设置属性
     * @param clazz
     * @return
     */
    public static ExcelMapping getFromXmlOrAnno(Class<?> clazz){
        List<ExcelProperty> propertyList = new ArrayList();
        if(!clazz.isAnnotationPresent(Excel.class))throw new AnnotationNotExistException(clazz.getName() +"注解不存在");
        Excel excel = clazz.getAnnotation(Excel.class);
        Field[] fields = clazz.getDeclaredFields();
        Arrays.stream(fields).forEach(o->{
            if(o.isAnnotationPresent(ExcelField.class)){
                ExcelProperty property = new ExcelProperty();
                ExcelField excelField = o.getAnnotation(ExcelField.class);
                property.setName(StringUtils.isNotBlank(excelField.value())?excelField.value():o.getName())
                        .setColumn(o.getName()).setComment(excelField.comment())
                        .setMaxLength(excelField.maxLength()).setWidth(excelField.width())
                        .setDateFormat(excelField.dateFormat()).setSort(excelField.sort()).setRequire(excelField.require())
                        .setExchange(excelField.exchange()).setMethodName(excelField.methodName()).setExchangeName(excelField.exchangeName())
                        .setCheckMethod(excelField.checkMethod())
                ;
                propertyList.add(property);
            }
        });
        if(propertyList.size()<=0)throw new AnnotationNotExistException(ExcelField.class.getName()+"注解不存在");
        /**
         * 设置列数排序
         */
        List<ExcelProperty> list = propertyList.stream().sorted(Comparator.comparing(ExcelProperty::getSort)).collect(toList());
        for (int i = 0; i < list.size(); i++) {list.get(i).setSort(i);}
        return new ExcelMapping().setPropertyList(list).setName(excel.value());
    }

    /**
     * 设置单元格数据类型
     * @param filed
     * @return
     * CELL_TYPE_NUMERIC 数值型 0
     * CELL_TYPE_STRING 字符串型 1
     * CELL_TYPE_FORMULA 公式型 2
     * CELL_TYPE_BLANK 空值 3
     * CELL_TYPE_BOOLEAN 布尔型 4
     * CELL_TYPE_ERROR 错误 5
     */
    public static CellType cellTypeMapping(Field filed){
        Class clazz = filed.getType();
        if(clazz==Number.class) return CellType.NUMERIC;
        if(clazz==String.class) return CellType.STRING;
        if(clazz==Boolean.class) return CellType.BOOLEAN;
        return CellType.STRING;
    }

}
