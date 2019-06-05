package com.git.poi.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author authorZhao
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelField {

  /**
   * 该字段对应列的名称
   * @return
   */
  String value() default "";

  /**
   * 复杂对象
   * user.username
   * @return
   */
  String name() default "";
  /**
   * 数据校验方法，电话邮箱等
   */
  String checkMethod() default "";
  /**
   * 是否需要转化
   */
  boolean exchange() default false;

  /**
   *转换的方法名，
   * brandName-> brnnd.class->getBrandIdName->getBrandId赋值给brandId
   * @return
   */
  String methodName() default "";

  /**
   * 转换后赋值给该字段
   * @return
   */
  String exchangeName() default"";

  /**
   * @return 单元格宽度[仅限表头] 默认-1(自动计算列宽)
   */
  short width() default -1;

  /**
   * @return 批注信息, 生成模板时生效
   */
  String comment() default "";

  /**
   * @return 最大长度, 读取时生效, 默认不限制
   */
  int maxLength() default -1;

  /**
   * 日期格式, 如: yyyy-MM-dd
   *
   * @return 日期格式
   */
  String dateFormat() default "yyyy-MM-dd";

  /**
   * 排序
   * @return
   */
  int sort() default 0;

  /**
   * 列的类型,默认字符串
   * NUMERIC(0,"数值型"),
   * STRING(1,"字符串型"),
   * FORMULA(2,"公式型"),
   * BLANK(3,"空值"),
   * BOOLEAN(4,"布尔型"),
   * ERROR(5,"错误");
   * @return
   */
  int type() default 1;

  /**
   * 是否为必填字段
   * @return
   */
  boolean require() default true;

}
