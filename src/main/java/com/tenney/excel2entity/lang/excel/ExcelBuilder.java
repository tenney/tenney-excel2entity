/**
 * 版权所有：tenney
 * 项目名称: eicsp
 * 类名称:ExcelBuilder.java
 * 包名称:com.tenney.excel2entity.lang.excel
 * 
 * 创建日期:2013年10月18日 下午6:55:48
 * 创建人： 唐雄飞		
 * <修改人>      <时间>      <版本号>    <描述>
 * 唐雄飞      2013年10月18日     	V1.0.0        N/A
 */

package com.tenney.excel2entity.lang.excel;

import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.tenney.excel2entity.ExcelConstants;
import com.tenney.excel2entity.support.GuideEntity;

/**
 * 类说明: <br/>
 * 创建人: 唐雄飞 <br/>
 * 创建日期:2013年10月18日 <br/>
 * 
 */
public class ExcelBuilder
{
    /**
     * 
     * 方法描述: 创建表格工作区 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月21日 <br>
     * @return <br>
     */
    public static Workbook getWorkBookInstance(String fileSuffix)
    {
        if(ExcelConstants.EXCEL_FILE_SUFFIX_XLS.equals(StringUtils.trimToEmpty(fileSuffix).toLowerCase())){
        	return new HSSFWorkbook();
        }
        return new XSSFWorkbook();
    }
    
    /**
     * 
     * 方法描述: 根据workbook类型生成富文本信息 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年11月4日 <br>
     * @param workbook
     * @param text
     * @return <br>
     */
    public static RichTextString getRichTextString(Workbook workbook,String text){
    	if(workbook instanceof XSSFWorkbook){
    		return new XSSFRichTextString(StringUtils.trimToEmpty(text));
    	}
    	return new HSSFRichTextString(StringUtils.trimToEmpty(text));
    }

    /**
     * 
     * 方法描述:  创建表格sheet工作簿 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月21日 <br>
     * @param workbook
     * @param entity
     * @return <br>
     */
    public static Sheet buildExcelSheet(Workbook workbook, GuideEntity entity)
    {
        return buildExcelSheet(workbook, entity.geteName());
    }

    public static Sheet buildExcelSheet(Workbook workbook, String sheetName)
    {
        Sheet sheet = workbook.createSheet(sheetName);
        sheet.setDefaultColumnWidth((short)15);
        /**
         * 前两个参数是要用来拆分的列数和行数。
         * 后两个参数是下面窗口的可见象限，
         * 第三个参数是右边区域可见的左边列数，
         * 第四个参数是下面区域可见的首行。
         */
     // 冻结第一行 
//        sheet.createFreezePane(0, 1, 0, 1);
        return sheet;
    }
    
    /**
     * 
     * 方法描述: 创建表格标题行的样式 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月21日 <br>
     * @param workbook
     * @return <br>
     */
    public static CellStyle buildTitleStyle(Workbook workbook){
        CellStyle style = workbook.createCellStyle();
        // 设置这些样式
        style.setFillForegroundColor(HSSFColorPredefined.TAN.getIndex());
//        style.setFillBackgroundColor(HSSFColor.TEAL.index);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        style.setBottomBorderColor(HSSFColorPredefined.GREY_50_PERCENT.getIndex());
        style.setLeftBorderColor(HSSFColorPredefined.GREY_50_PERCENT.getIndex());
        style.setRightBorderColor(HSSFColorPredefined.GREY_50_PERCENT.getIndex());
        style.setTopBorderColor(HSSFColorPredefined.GREY_50_PERCENT.getIndex());
        
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setAlignment(HorizontalAlignment.CENTER);
        
        Font font = workbook.createFont();
        font.setColor(IndexedColors.ORANGE.getIndex());
//        if(workbook instanceof XSSFWorkbook){
//        	font.setColor(IndexedColors.BLUE.getIndex());
//        }else{
//        	font.setColor(HSSFColor.BLUE.index);
//        }
        font.setFontHeightInPoints((short) 10);
        font.setBold(true);
        // 把字体应用到当前的样式
        style.setFont(font);
        
        return style;
    }
    
    /**
     * 
     * 方法描述: 创建表格标题行的样式 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月21日 <br>
     * @param workbook
     * @return <br>
     */
    public static CellStyle buildTitlenameStyle(Workbook workbook){
        CellStyle style = workbook.createCellStyle();
        // 设置这些样式
        style.setFillForegroundColor((short) 13);
//        style.setFillBackgroundColor(HSSFColor.TEAL.index);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setWrapText(true);
        
        style.setBottomBorderColor(HSSFColorPredefined.GREY_50_PERCENT.getIndex());
        style.setLeftBorderColor(HSSFColorPredefined.GREY_50_PERCENT.getIndex());
        style.setRightBorderColor(HSSFColorPredefined.GREY_50_PERCENT.getIndex());
        style.setTopBorderColor(HSSFColorPredefined.GREY_50_PERCENT.getIndex());
        
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);

        Font font = workbook.createFont();
        font.setColor(IndexedColors.BLUE.getIndex());
//        if(workbook instanceof XSSFWorkbook){
//        	font.setColor(IndexedColors.BLUE.getIndex());
//        }else{
//        	font.setColor(HSSFColor.BLUE.index);
//        }
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        // 把字体应用到当前的样式
        style.setFont(font);
        
        return style;
    }
    
    
    
    /**
     * 
     * 方法描述: 创建错误行的样式<br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月25日 <br>
     * @param workbook
     * @return <br>
     */
    public static CellStyle buildErrorStyle(Workbook workbook){
        CellStyle errorStyle = workbook.createCellStyle();
        errorStyle.setFillForegroundColor(HSSFColorPredefined.YELLOW.getIndex());
        errorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = workbook.createFont();
//        font.setColor(HSSFColor.RED.index);
        if(workbook instanceof XSSFWorkbook){
        	font.setColor(IndexedColors.BLUE.getIndex());
        }else{
        	font.setColor(HSSFColorPredefined.BLUE.getIndex());
        }
        errorStyle.setFont(font);
        return errorStyle;
    }
    
    /**
     * 
     * 方法描述: 创建提示消息样式 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月25日 <br>
     * @param workbook
     * @return <br>
     */
    public static CellStyle buildMessageStyle(Workbook workbook){
        CellStyle msgStyle = workbook.createCellStyle();
        msgStyle.setFillForegroundColor(HSSFColorPredefined.LIME.getIndex());
        msgStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = workbook.createFont();
        if(workbook instanceof XSSFWorkbook){
        	font.setColor(IndexedColors.BLUE.getIndex());
        }else{
        	font.setColor(HSSFColorPredefined.BLUE.getIndex());
        }
        msgStyle.setFont(font);
        return msgStyle;
    }
    
    /**
     * 
     * 方法描述: 根据文件流创建Excel的HSSFWorkbook <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月24日 <br>
     * @param input
     * @return
     * @throws IOException <br>
     * @throws InvalidFormatException 
     */
    public static Workbook readWorkbook(InputStream input) throws IOException, InvalidFormatException{
//        return new HSSFWorkbook(new POIFSFileSystem(input));
    	
    	if (!input.markSupported()) {  
    		input = new PushbackInputStream(input, 8);  
    	} 
    	if(FileMagic.OLE2 == FileMagic.valueOf(input)) {
//        if (POIFSFileSystem.hasPOIFSHeader(input)) {
            return new HSSFWorkbook(input);
        }
        if(FileMagic.OOXML == FileMagic.valueOf(input)) {
//        if (POIXMLDocument.hasOOXMLHeader(input)) {
            return new XSSFWorkbook(OPCPackage.open(input));
        }
        throw new IllegalArgumentException("无法识别的Excel文档.");
    }

    /**
     * 根据表格列类型取值
     * @param cell
     * @return
     */
    public static Object getCellValue(Cell cell){
        if(cell == null){
            return null;
        }
        Object cellValue = null;
        switch(cell.getCellType()){
            case Cell.CELL_TYPE_NUMERIC:
                /**
                 * 所有日期格式都可以通过getDataFormat()值来判断，用于取自定义日期格式
                 yyyy-MM-dd------14
                 yyyy年m月d日-----31
                 yyyy年m月--------57
                 m月d日  ----------58
                 HH:mm-----------20
                 h时mm分  --------32
                 */
                int format = cell.getCellStyle().getDataFormat();
                if (DateUtil.isCellDateFormatted(cell) || format == 14 || format == 31 || format == 57 || format == 58 ||format == 20 || format == 32) {
//                    cellValue = cell.getDateCellValue();
                    cellValue = DateUtil.getJavaDate(cell.getNumericCellValue());
                }else{
                	cellValue = cell.getNumericCellValue();
//                    double value = cell.getNumericCellValue();
//                    DecimalFormat dformat = new DecimalFormat();
//                    // 单元格设置成常规
//                    if (cell.getCellStyle().getDataFormatString().equals("General")) {
//                        dformat.applyPattern("#");
//                    }
//                    cellValue = dformat.format(value);
                }
                break;
            case Cell.CELL_TYPE_STRING:
//                cellValue = cell.getStringCellValue();
                cellValue = cell.getRichStringCellValue().getString();
                break;
            case Cell.CELL_TYPE_FORMULA:
                cellValue=String.valueOf(cell.getCellFormula());
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                cellValue=String.valueOf(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_ERROR:
                cellValue=String.valueOf(cell.getErrorCellValue());
                break;
        }
        return cellValue;
    }
    
    
    
	public static DataValidation setDataValidation(Sheet sheet, String[] textList, int firstRow, int endRow, int firstCol, int endCol) {

        DataValidationHelper helper = sheet.getDataValidationHelper();
        //加载下拉列表内容
        DataValidationConstraint constraint = helper.createExplicitListConstraint(textList);
        constraint.setExplicitListValues(textList);
        
        //设置数据有效性加载在哪个单元格上。四个参数分别是：起始行、终止行、起始列、终止列
        CellRangeAddressList regions = new CellRangeAddressList((short) firstRow, (short) endRow, (short) firstCol, (short) endCol);
    
        //数据有效性对象
        DataValidation data_validation = helper.createValidation(constraint, regions);
    
        return data_validation;
    }
}
