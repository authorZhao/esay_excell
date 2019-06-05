package com.git.poi.util;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * @author authorZhao
 */
public class PoIUtil {

  private static final int mDefaultRowAccessWindowSize = 100;


  private static SXSSFWorkbook newSXSSFWorkbook(int rowAccessWindowSize) {
    return new SXSSFWorkbook(rowAccessWindowSize);
  }

  public static SXSSFWorkbook newSXSSFWorkbook() {
    return PoIUtil.newSXSSFWorkbook(PoIUtil.mDefaultRowAccessWindowSize);
  }

  public static Sheet newSXSSFSheet(SXSSFWorkbook wb, String sheetName) {
    return wb.createSheet(sheetName);
  }

  public static Row newSXSSFRow(SXSSFSheet sheet, int index) {
    return sheet.createRow(index);
  }

  public static Cell newSXSSFCell(SXSSFRow row, int index) {
    return row.createCell(index);
  }

  public static void setColumnWidth(
      SXSSFSheet sheet, int index, Short width, String value) {
    boolean widthNotHaveConfig = (null == width || width == -1);
    if (widthNotHaveConfig && !ValidatorUtil.isEmpty(value)) {
      sheet.setColumnWidth(index, (short) (value.length() * 2048));
    } else {
      width = widthNotHaveConfig ? 200 : width;
      sheet.setColumnWidth(index, (short) (width * 35.7));
    }
  }

 /* public static void setColumnCellRange(SXSSFSheet sheet, Options options,
      int firstRow, int endRow,
      int firstCell, int endCell) {
    if (null != options) {
      String[] datasource = options.get();
      if (null != datasource && datasource.length > 0) {
        if (datasource.length > 100) {
          throw new RuntimeException("Options item too much.");
        }

        DataValidationHelper validationHelper = sheet.getDataValidationHelper();
        DataValidationConstraint explicitListConstraint = validationHelper
            .createExplicitListConstraint(datasource);
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCell,
            endCell);
        DataValidation validation = validationHelper
            .createValidation(explicitListConstraint, regions);
        validation.setSuppressDropDownArrow(true);
        validation.createErrorBox("提示", "请从下拉列表选取");
        validation.setShowErrorBox(true);
        sheet.addValidationData(validation);
      }
    }
  }*/

  public static void write(SXSSFWorkbook wb, OutputStream out) {
    try {
      if (null != out) {
        wb.write(out);
        out.flush();
      }
    } catch (IOException e) {
      e.printStackTrace();
    } finally {
      if (null != out) {
        try {
          out.close();
        } catch (IOException e) {
          e.printStackTrace();
        }
      }
    }
  }

  public static void download(
      SXSSFWorkbook wb, HttpServletResponse response, String filename) {
    try {
      OutputStream out = response.getOutputStream();
      response.setContentType(Const.XLSX_CONTENT_TYPE);
      response.setHeader(Const.XLSX_HEADER_KEY,
          String.format(Const.XLSX_HEADER_VALUE_TEMPLATE, filename));
      PoIUtil.write(wb, out);
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  public static Object convertByExp(Object propertyValue, String converterExp)
      throws Exception {
    try {
      String[] convertSource = converterExp.split(",");
      for (String item : convertSource) {
        String[] itemArray = item.split("=");
        if (itemArray[0].equals(propertyValue)) {
          return itemArray[1];
        }
      }
    } catch (Exception e) {
      throw e;
    }
    return propertyValue;
  }

  public static int countNullCell(String ref, String ref2) {
    // excel2007最大行数是1048576，最大列数是16384，最后一列列名是XFD
    String xfd = ref.replaceAll("\\d+", "");
    String xfd_1 = ref2.replaceAll("\\d+", "");

    xfd = PoIUtil.fillChar(xfd, 3, '@', true);
    xfd_1 = PoIUtil.fillChar(xfd_1, 3, '@', true);

    char[] letter = xfd.toCharArray();
    char[] letter_1 = xfd_1.toCharArray();
    int res =
        (letter[0] - letter_1[0]) * 26 * 26 + (letter[1] - letter_1[1]) * 26 + (letter[2]
            - letter_1[2]);
    return res - 1;
  }

  private static String fillChar(String str, int len, char let, boolean isPre) {
    int len_1 = str.length();
    if (len_1 < len) {
      if (isPre) {
        for (int i = 0; i < (len - len_1); i++) {
          str = let + str;
        }
      } else {
        for (int i = 0; i < (len - len_1); i++) {
          str = str + let;
        }
      }
    }
    return str;
  }

  public static void checkExcelFile(File file) {
    String filename = null != file ? file.getAbsolutePath() : null;
    if (null == filename || !file.exists()) {
      throw new RuntimeException("Excel file[" + filename + "] does not exist.");
    }
    if (!filename.endsWith(Const.XLSX_SUFFIX)) {
      throw new RuntimeException(
          "[" + filename + "]Only .xlsx formatted files are supported.");
    }
  }

  /**
   * ；两行复制
   * @param row
   * @param row1
   */
  public static void copyPoiRow(Row row, Row row1) {
    for (int i = 0; i < row.getLastCellNum() ; i++) {
      Cell cell0 = row.getCell(i);
      if(cell0==null)cell0=row.createCell(i);
      Cell cell1 = row1.createCell(i);
      int type = cell0.getCellType();
      cell1.setCellType(cell0.getCellType());
      switch (type){
        case 0:cell1.setCellValue(cell0.getNumericCellValue());break;
        case 1:cell1.setCellValue(cell0.getStringCellValue());break;
        case 2:cell1.setCellValue(cell0.getCellFormula());break;
        case 4:cell1.setCellValue(cell0.getBooleanCellValue());break;
        default: cell1.setCellValue("");break;
      }
    }
  }


}
