package com.git.poi.factory;


import com.git.poi.validator.Options;

/**
 * @author authorZhao
 */

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

  public String getName() {
    return name;
  }

  public ExcelProperty setName(String name) {
    this.name = name;
    return this;
  }

  public String getColumn() {
    return column;
  }

  public ExcelProperty setColumn(String column) {
    this.column = column;
    return this;
  }

  public Short getWidth() {
    return width;
  }

  public ExcelProperty setWidth(Short width) {
    this.width = width;
    return this;
  }

  public String getComment() {
    return comment;
  }

  public ExcelProperty setComment(String comment) {
    this.comment = comment;
    return this;
  }

  public Integer getMaxLength() {
    return maxLength;
  }

  public ExcelProperty setMaxLength(Integer maxLength) {
    this.maxLength = maxLength;
    return this;
  }

  public String getDateFormat() {
    return dateFormat;
  }

  public ExcelProperty setDateFormat(String dateFormat) {
    this.dateFormat = dateFormat;
    return this;
  }

  public String getWriteConverter() {
    return writeConverter;
  }

  public ExcelProperty setWriteConverter(String writeConverter) {
    this.writeConverter = writeConverter;
    return this;
  }

  public String getReadConverter() {
    return readConverter;
  }

  public ExcelProperty setReadConverter(String readConverter) {
    this.readConverter = readConverter;
    return this;
  }

  public Options getOptions() {
    return options;
  }

  public ExcelProperty setOptions(Options options) {
    this.options = options;
    return this;
  }

  public Integer getSort() {
    return sort;
  }

  public ExcelProperty setSort(Integer sort) {
    this.sort = sort;
    return this;
  }

  public Integer getCellType() {
    return cellType;
  }

  public ExcelProperty setCellType(Integer cellType) {
    this.cellType = cellType;
    return this;
  }

  public String getReason() {
    return reason;
  }

  public ExcelProperty setReason(String reason) {
    this.reason = reason;
    return this;
  }

  public Boolean getRequire() {
    return require;
  }

  public ExcelProperty setRequire(Boolean require) {
    this.require = require;
    return this;
  }

  public String getMethodName() {
    return methodName;
  }

  public ExcelProperty setMethodName(String methodName) {
    this.methodName = methodName;
    return this;
  }

  public String getExchangeName() {
    return exchangeName;
  }

  public ExcelProperty setExchangeName(String exchangeName) {
    this.exchangeName = exchangeName;
    return this;
  }

  public Boolean getExchange() {
    return exchange;
  }

  public ExcelProperty setExchange(Boolean exchange) {
    this.exchange = exchange;
    return this;
  }

  public String getCheckMethod() {
    return checkMethod;
  }

  public ExcelProperty setCheckMethod(String checkMethod) {
    this.checkMethod = checkMethod;
    return this;
  }
}
