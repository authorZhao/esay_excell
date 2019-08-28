package com.git.poi.util;

import com.git.poi.mapping.ExcelProperty;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.util.List;

public class WriteUtil {

    public static XSSFSheet setExcellXlsxHeader(XSSFWorkbook workbook, List<ExcelProperty> propertyList, String sheetName){
        XSSFSheet sheet = workbook.createSheet(sheetName);
        //宽度自适应
        sheet.autoSizeColumn(0, true);
        XSSFCellStyle style = getHeaderCellStyle(workbook);
        Row row = sheet.createRow(0);
        XSSFDrawing draw = sheet.createDrawingPatriarch();
        for (int i = 0; i < propertyList.size() ; i++) {
            sheet.setColumnWidth(i, (int)(propertyList.get(i).getName().getBytes().length * 1.2d * 256 > 12 * 256 ? propertyList.get(i).getName().getBytes().length * 1.2d * 256 : 12 * 256));
            Cell cell = row.createCell(i);
            cell.setCellValue(propertyList.get(i).getName());
            cell.setCellStyle(style);
            ClientAnchor clientAnchor = new XSSFClientAnchor(0,0,0,0,i,1,i+1,5);
            XSSFComment comment = draw.createCellComment(clientAnchor);
            comment.setString(new XSSFRichTextString(propertyList.get(i).getComment()));
            comment.setVisible(false);
            //添加作者,选中B5单元格,看状态栏
            comment.setAuthor("authorZhao");
            cell.setCellComment(comment);
        }
        return sheet;
    }

    private static XSSFCellStyle getHeaderCellStyle(XSSFWorkbook workbook) {
        XSSFCellStyle xlsxStyle = null;
        if (xlsxStyle == null) {
            xlsxStyle = workbook.createCellStyle();
            XSSFFont font = workbook.createFont();
            xlsxStyle.setFillForegroundColor((short) 12);
            xlsxStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);//填充方式
            xlsxStyle.setBorderTop(BorderStyle.MEDIUM);
            xlsxStyle.setBorderRight(BorderStyle.MEDIUM_DASH_DOT);
            xlsxStyle.setBorderBottom(BorderStyle.DOUBLE);
            xlsxStyle.setBorderLeft(BorderStyle.THICK);
            xlsxStyle.setAlignment(HorizontalAlignment.LEFT);// 对齐
            xlsxStyle.setFillBackgroundColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
            xlsxStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
            font.setColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
            // 应用标题字体到标题样式
            xlsxStyle.setFont(font);//设置字体
            //设置单元格文本形式
            DataFormat dataFormat = workbook.createDataFormat();
            xlsxStyle.setDataFormat(dataFormat.getFormat("@"));
        }
        return xlsxStyle;
    }


}
