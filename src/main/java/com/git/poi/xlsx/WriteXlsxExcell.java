package com.git.poi.xlsx;

import com.git.poi.mapping.ExcelMapping;
import com.git.poi.mapping.ExcelProperty;
import com.git.poi.mapping.FailRecord;
import com.git.poi.util.PoiUtil;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Method;
import java.util.List;

/**
 * 导出数据
 * @author authorZhao
 */
public class WriteXlsxExcell {
    public static Logger logger = LoggerFactory.getLogger(WriteXlsxExcell.class);

    private final ExcelMapping mExcelMapping;
    private final Integer mMaxSheetRecords;

    public WriteXlsxExcell(ExcelMapping excelMapping, Integer maxSheetRecords) {
        mExcelMapping = excelMapping;
        mMaxSheetRecords = maxSheetRecords;
    }
    public Workbook generateXlsxWorkbook(List<?> data, boolean isTemplate)  {
        return generateXlsxWorkbook(data,isTemplate,null);
    }

    /**
     *
     * @param data 导出的数据
     * @param isTemplate 是不是模板，是的话只有两行
     * @param template 模板的规则文字
     * @return
     */
    public Workbook generateXlsxWorkbook(List<?> data, boolean isTemplate, String template)  {
        //不用SXSSFWorkbook，太坑了
        //Workbook workbook = new SXSSFWorkbook();
        Workbook workbook = new XSSFWorkbook();
        List<ExcelProperty> propertyList = mExcelMapping.getPropertyList();
        Double sheetNo =1.0;
        if(data!=null&&data.size()>0)sheetNo = Math.ceil(data.size() / (double) mMaxSheetRecords);
        for (int index = 0; index <= (sheetNo == 0.0 ? sheetNo : sheetNo - 1); index++) {
            String sheetName = mExcelMapping.getName() + (index == 0 ? "" : "_" + index);
            //设置表头
            Sheet sheet = setXlsxHeader(workbook, propertyList, sheetName, isTemplate);
            //数据校验
            //setValidation(mExcelMapping.getPropertyList().get(3),sheet,new String[]{"是","否"},3);
            if(isTemplate)continue;
            //设置表身

            if (null != data && data.size() > 0) {
                int startNo = index * mMaxSheetRecords;
                int endNo = Math.min(startNo + mMaxSheetRecords, data.size());
                //从前两个对象里面取出，可能都不存在
                Class clazz = data.get(0).getClass()==null?data.get(0).getClass():data.get(1).getClass();
                for (int i = startNo; i < endNo; i++) {
                    Row bodyRow = sheet.createRow(i-startNo+1);
                    for (int j = 0; j < propertyList.size(); j++) {
                        Cell cell = bodyRow.createCell(j);
                        //设置值
                        Object obj = null;
                        String methodName = "get"+(new StringBuilder()).append(Character.toUpperCase(propertyList.get(j).getColumn().charAt(0))).append(propertyList.get(j).getColumn().substring(1)).toString();
                        try{
                            Method m2 = clazz.getDeclaredMethod(methodName);
                            obj = m2.invoke(data.get(i));
                        } catch (Exception e) {
                            logger.error("反射处理失败");
                            e.printStackTrace();
                        }
                        if(obj instanceof String)cell.setCellValue((String)obj);
                        if(obj instanceof Number){
                            if(obj instanceof Long)cell.setCellValue((Long)obj);
                            if(obj instanceof Double)cell.setCellValue((Double)obj);
                            if(obj instanceof Integer)cell.setCellValue((Integer)obj);
                        }
                    }
                }
            }
        }
        return workbook;
    }

    /**
     * 根据失败的行数返回原表格
     * @param data
     * @return
     */
    public Workbook generateFailXlsxWorkbook(List<FailRecord> data, File file, String name)  {
        Workbook workbook = null;
        Workbook newBook = null;
        try {
            newBook = new SXSSFWorkbook();
            workbook = new XSSFWorkbook(new FileInputStream(file));
        } catch (IOException e) {
            e.printStackTrace();
        }
        Sheet newSheet = setXlsxHeader(newBook,mExcelMapping.getPropertyList(),name,false);
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        Row newRow0 = newSheet.createRow(0);
        PoiUtil.copyPoiRow(row,newRow0);
        newRow0.createCell(row.getLastCellNum()).setCellValue("失败原因");
        for (int i = 0; i < data.size() ; i++) {
            Row newRow = newSheet.createRow(i+1);
            PoiUtil.copyPoiRow(sheet.getRow(data.get(i).getRowNum()),newRow);
            newRow.createCell(row.getLastCellNum()).setCellValue(data.get(i).getMessage()+"   位于原表格"+data.get(i).getRowNum()+"行");
        }
        logger.info("失败数据生成完毕");
        return newBook;
    }


    /**
     *
     * @param data 验证的数据源
     * @param sheet 表格
     * @param regex 验证规则
     * @param clo 验证的列数
     */
    public void setValidation(List<?> data, Sheet sheet, String[] regex, int clo) {
        DataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) sheet);
        DataValidationConstraint dvConstraint = dvHelper
                .createExplicitListConstraint(regex);
        // 设置区域边界
        CellRangeAddressList addressList = new CellRangeAddressList(1, 50000, clo, clo);
        DataValidation validation =  dvHelper.createValidation(dvConstraint, addressList);
        // 输入非法数据时，弹窗警告框
        validation.setShowErrorBox(true);
        // 设置提示框
        String test="";
        if(regex.length<=10){
            test = PoiUtil.getString(regex);
            test = "请在（"+test+"）之间选择";
        }else{
            test = "请在合法范围内输入";
        }
        validation.createPromptBox("温馨提示",test );
        validation.setShowPromptBox(true);
        validation.setEmptyCellAllowed(true);
        validation.setSuppressDropDownArrow(true);
        validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        sheet.addValidationData(validation);

    }

    /**
     * 表头设置
     * @param workbook 表
     * @param propertyList 属性结合
     * @param sheetName 表名
     * @param isTemplate 是否是模板
     * @return
     */
    private Sheet setXlsxHeader(Workbook workbook, List<ExcelProperty> propertyList, String sheetName, boolean isTemplate) {
        Sheet sheet = workbook.createSheet(sheetName);
        //sheet.setColumnWidth(0, 1000);
        //自适应
        sheet.autoSizeColumn(0, true);
        CellStyle style = getHeaderCellStyle(workbook);
        Row row = sheet.createRow(0);
        XSSFDrawing draw = (XSSFDrawing)sheet.createDrawingPatriarch();
        for (int i = 0; i < propertyList.size() ; i++) {
            sheet.setColumnWidth(i, (int)(propertyList.get(i).getName().getBytes().length * 1.2d * 256 > 12 * 256 ? propertyList.get(i).getName().getBytes().length * 1.2d * 256 : 12 * 256));
            //sheet.setDefaultColumnStyle(i,style);
            Cell cell = row.createCell(i);

            cell.setCellValue(propertyList.get(i).getName());
            cell.setCellStyle(style);
            //获取批注对象
            //(int dx1, int dy1, int dx2, int dy2, short col1, int row1, short col2, int row2)
            //前四个参数是坐标点,后四个参数是编辑和显示批注时的大小.
            /**
             *   *@param dx1  the x coordinate in EMU within the first cell. dx1 第1个单元格中x轴的偏移量
             *   * @param dy1  the y coordinate in EMU within the first cell. dy1 第1个单元格中y轴的偏移量
             *   * @param dx2  the x coordinate in EMU within the second cell. dx2 第2个单元格中x轴的偏移量
             *   * @param dy2  the y coordinate in EMU within the second cell. dy2 第2个单元格中y轴的偏移量
             *   * @param col1 the column (0 based) of the first cell.col1 第1个单元格的列号
             *   * @param row1 the row (0 based) of the first cell. row1 第1个单元格的行号
             *   * @param col2 the column (0 based) of the second cell.col2 第2个单元格的列号
             *   * @param row2 the row (0 based) of the second cell. row2 第2个单元格的行号
             *   * @return the newly created client anchor
             */
            ClientAnchor clientAnchor=draw.createAnchor(0,0,0,0,i,1,i+1,5);
            //输入批注信息
            Comment comment = draw.createCellComment(clientAnchor);
            comment.setString(new XSSFRichTextString(propertyList.get(i).getComment()));
            comment.setVisible(false);
            //添加作者,选中B5单元格,看状态栏
            comment.setAuthor("authorZhao");
            //将批注添加到单元格对象中
            cell.setCellComment(comment);
        }
        //如果是模板，设置规则文字
       /* if(isTemplate){
            CellRangeAddress region = new CellRangeAddress(1, 1, 0, propertyList.size()-1);
            Row row1 = sheet.createRow(1);
            row1.createCell(0).setCellValue(template);
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFillBackgroundColor(IndexedColors.SEA_GREEN.getIndex());
            cellStyle.setFillForegroundColor(IndexedColors.SEA_GREEN.getIndex());
            cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
            cellStyle.setWrapText(true);
            row1.getCell(0).setCellStyle(cellStyle);
            System.out.println(row.getHeight());
            row1.setHeight((short)(row.getHeight()*15));
            sheet.addMergedRegion(region);
        }*/
        for (int i=0;i<row.getLastCellNum();i++) {
            Comment comment = row.getCell(i).getCellComment();
        }
        logger.info("表头设置完毕");

        return sheet;
    }


    /**
     * 表头样式
     * @param wb
     * @return
     */
    public CellStyle getHeaderCellStyle(Workbook wb) {
        CellStyle mHeaderCellStyle = null;
        if (null == mHeaderCellStyle) {
            mHeaderCellStyle = wb.createCellStyle();
            Font font = wb.createFont();
            mHeaderCellStyle.setFillForegroundColor((short) 12);
            mHeaderCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);//填充方式
            mHeaderCellStyle.setBorderTop(BorderStyle.MEDIUM);
            mHeaderCellStyle.setBorderRight(BorderStyle.MEDIUM_DASH_DOT);
            mHeaderCellStyle.setBorderBottom(BorderStyle.DOUBLE);
            mHeaderCellStyle.setBorderLeft(BorderStyle.THICK);
            mHeaderCellStyle.setAlignment(HorizontalAlignment.LEFT);// 对齐
            mHeaderCellStyle.setFillBackgroundColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
            mHeaderCellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
            font.setColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
            // 应用标题字体到标题样式
            mHeaderCellStyle.setFont(font);//设置字体
            //设置单元格文本形式
            DataFormat dataFormat = wb.createDataFormat();
            mHeaderCellStyle.setDataFormat(dataFormat.getFormat("@"));
        }
        return mHeaderCellStyle;
    }
/**
 * 在使用NPOI技术开发自动操作EXCEL软件时遇到不能精确设置列宽的问题。
 *
 * 如
 * ISheet sheet1 = hssfworkbook.CreateSheet("Sheet1");
 * sheet1.SetColumnWidth(0,  50 * 256);   // 在EXCEL文档中实际列宽为49.29
 * sheet1.SetColumnWidth(1, 100 * 256);   // 在EXCEL文档中实际列宽为99.29
 * sheet1.SetColumnWidth(2, 150 * 256);   // 在EXCEL文档中实际列宽为149.29
 *
 * 到此一般人应该知道问题出在哪了，解决方法如下:
 * ISheet sheet1 = hssfworkbook.CreateSheet("Sheet1");
 * sheet1.SetColumnWidth(0,  (int)((50 + 0.72) * 256));   // 在EXCEL文档中实际列宽为50
 * sheet1.SetColumnWidth(1,  (int)((100 + 0.72) * 256));   // 在EXCEL文档中实际列宽为100
 * sheet1.SetColumnWidth(2,  (int)((150 + 0.72) * 256));   // 在EXCEL文档中实际列宽为150
 */
}
