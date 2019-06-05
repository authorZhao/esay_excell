package com.git.poi.factory;

import java.util.List;

public class ExcelMapping {

  private String name;
  private List<ExcelProperty> propertyList;

  public String getName() {
    return name;
  }

  public ExcelMapping setName(String name) {
    this.name = name;
    return this;
  }

  public List<ExcelProperty> getPropertyList() {
    return propertyList;
  }

  public ExcelMapping setPropertyList(List<ExcelProperty> propertyList) {
    this.propertyList = propertyList;
    return this;
  }
}
