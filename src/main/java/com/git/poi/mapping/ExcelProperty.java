package com.git.poi.mapping;


import com.git.poi.validator.Options;
import lombok.Data;
import lombok.experimental.Accessors;

/**
 * @author authorZhao
 */
@Data
@Accessors(chain = true)
public class ExcelProperty {
  /**
   * name，列的名字
   */
  private String name;
  /**
   * 字段，
   */
  private String column;
  /**
   * 宽度
   */
  private Short width;
  /**
   * 注释
   */
  private String comment;
    /**
     * 最大长度
     */
  private Integer maxLength;
    /**
     * 日期格式
     */
  private String dateFormat;
  /**
   * 写转换器
   */
  private String writeConverter;
  /**
   * 读转换器
   */
  private String readConverter;

  /**
   * 校验规则
   */
  private Options options;

  /**
   * 排序，指定该字段在第几列
   */
  private Integer sort;
  /**
   * 字段类型
   */
  private Integer cellType;
  /**
   * 失败原因
   */
  private String reason;

  /**
   * 是否为必填字段
   * @return
   */
  private Boolean require;

  /**
   *转换的方法名，
   * brandName-> brnnd.class->getBrandIdName->getBrandId赋值给brandId
   * @return
   */
  private String methodName;

  /**
   * 转换后赋值给该字段
   * @return
   */
  private String exchangeName;

  /**
   * 是否需要转化
   */
  private Boolean exchange;

  /**
   * 校验的方法名
   */
  private String checkMethod;

}
