package com.git.poi.validator;

import com.git.poi.exception.ColumnConvertException;
import com.git.poi.mapping.ExcelProperty;

import java.lang.reflect.Method;
import java.util.List;

/**
 * 数据转化，目前只支持字符串类型的
 * 重新确认，将分类字符串转换为分类id
 * @author authorZhao
 */

public class Options {

    /**
     *
     * @param list 分类的id集合
     * @param excelProperty
     * @param obj 需要转换的对象
     * @throws Exception
     */
    public static void getString(List list , ExcelProperty excelProperty, Object obj) throws Exception {
        if(list==null||list.size()<=0)return;
        Boolean exchage = excelProperty.getExchange();
        if(exchage==null||!exchage)return;
        Class clazz = list.get(0).getClass();
        try {
            //List<Goods> goods.getName()
            Method m1 = clazz.getMethod(excelProperty.getMethodName());
            //String methodName = "set" + (new StringBuilder()).append(Character.toUpperCase(excelProperty.getColumn().charAt(0))).append(excelProperty.getColumn().substring(1)).toString();
            //goodsXlsx.getName()
            Method m2 = obj.getClass().getMethod(excelProperty.getMethodName());
            String value2 = (String)m2.invoke(obj);
            //goodsXlsx.setId()
            Method m3 = obj.getClass().getMethod(excelProperty.getExchangeName(),Integer.class);

            Object object = list.stream().filter(o-> {
                Integer value1 = null;
                try {
                    value1 = (Integer)m1.invoke(o);
                } catch (Exception e) {
                   throw new ColumnConvertException();
                }
                return value1.equals(value2);
            }).findFirst().orElse(null);
            m3.invoke(obj,object);
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        }
    }


}
