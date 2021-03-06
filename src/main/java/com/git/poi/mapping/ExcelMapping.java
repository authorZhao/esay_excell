package com.git.poi.mapping;


import lombok.Data;
import lombok.experimental.Accessors;

import java.util.List;
@Data
@Accessors(chain = true)
public class ExcelMapping {
  private String name;
  private List<ExcelProperty> propertyList;
}
