/**
 * 版权所有：tenney
 * 项目名称: eicsp
 * 类名称:ExcelExportToExcel.java
 * 包名称:com.tenney.excel2entity
 * 
 * 创建日期:2013年10月18日 下午7:15:35
 * 创建人： 唐雄飞		
 * <修改人>      <时间>      <版本号>    <描述>
 * 唐雄飞      2013年10月18日     	V1.0.0        N/A
 */

package com.tenney.excel2entity;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.Arrays;
import java.util.Collection;
import java.util.Date;
import java.util.Map;

import javax.imageio.ImageIO;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.tenney.excel2entity.ExcelConstants.ImageType;
import com.tenney.excel2entity.lang.DateHelper;
import com.tenney.excel2entity.lang.ExcelGuideException;
import com.tenney.excel2entity.lang.excel.ExcelBuilder;
import com.tenney.excel2entity.support.GuideEntity;
import com.tenney.excel2entity.support.GuideEntityField;
import com.tenney.excel2entity.support.GuideTitle;
import com.tenney.excel2entity.support.IExcelEntity;

/**
 * 类说明: <br/>
 * 创建人: 唐雄飞 <br/>
 * 创建日期:2013年10月18日 <br/>
 * 
 */
public class ExcelExportToExcel
{
    private static final Logger logger = Logger.getLogger(ExcelExportToExcel.class);
    
    private GuideEntity entity;
    private Workbook workbook;

    /**
     * ExcelExportToExcel.java的构造函数
     */
    public ExcelExportToExcel()
    {
        super();
    }

    /**
     * ExcelExportToExcel.java的构造函数
     * 
     * @param entity
     * @param workbook
     */
    public ExcelExportToExcel(GuideEntity entity, Workbook workbook)
    {
        super();
        this.entity = entity;
        this.workbook = workbook;
    }
    
    @SuppressWarnings("rawtypes")
    public boolean invokeExport(Collection<?> dataSet) throws IllegalAccessException, InvocationTargetException, NoSuchMethodException ,ExcelGuideException{
        boolean flag = true;
        if(null != this.workbook && null != this.entity){
            //创建sheet
            Sheet sheet = ExcelBuilder.buildExcelSheet(workbook, entity);
            CellStyle titleStyle = ExcelBuilder.buildTitleStyle(workbook);
            CellStyle titlenameStyle = ExcelBuilder.buildTitlenameStyle(workbook);

            //填充标题 
            Row titleRow = sheet.createRow(ExcelConstants.TITLE_ROW_IDX);
            GuideTitle title = this.entity.getTitle();
            System.out.println(title);
            short idx = 0;
            if(title != null){
            	CellRangeAddress cra=new CellRangeAddress(0, title.getRowspan(), 0, title.getColspan()); 
            	sheet.addMergedRegion(cra);  
            	Cell cell_1 = titleRow.createCell(0);  
            	cell_1.setCellStyle(titlenameStyle);
            	cell_1.setCellValue(title.getExcelTitle());
            	
            	Row titleRow1 = sheet.createRow(title.getRowspan() + 1);
            	for(GuideEntityField field : this.entity.getFields()){
                    Cell cell = titleRow1.createCell(idx++);
                    cell.setCellStyle(titleStyle);
                    if(!field.getImported()){
//                    	cell.setCellComment();
                    }
//                    HSSFRichTextString fieldText = new HSSFRichTextString(field.getExcelTitle());
//                    cell.setCellValue(fieldText);
                    cell.setCellValue(field.getExcelTitle());
                }
            }else{
            	 for(GuideEntityField field : this.entity.getFields()){
                     Cell cell = titleRow.createCell(idx++);
                     cell.setCellStyle(titleStyle);
                     if(!field.getImported()){
//                     	cell.setCellComment();
                     }
//                     HSSFRichTextString fieldText = new HSSFRichTextString(field.getExcelTitle());
//                     cell.setCellValue(fieldText);
                     cell.setCellValue(field.getExcelTitle());
                 }
            }
            
            //填充数据内容
            if(dataSet != null && !dataSet.isEmpty()){
				int rowIndex = ExcelConstants.TITLE_ROW_IDX +1;
				if(title != null){
					rowIndex = title.getRowspan() + 2;
				}
               
                for(Object data:dataSet){
                    short cell = 0;
                    //创建行，在标题行下面
                    Row dataRow = sheet.createRow(rowIndex++);
                  //行高不为空
                	if(this.entity.getHeightOfRows() != null){ //只在第列
                		dataRow.setHeightInPoints(this.entity.getHeightOfRows());
                	}
                    
                    Drawing patriarch = sheet.createDrawingPatriarch();//一定要放在循环外,只能声明一次。
                    //填充列数据
                    for(GuideEntityField field : this.entity.getFields()){
                        Cell dataCell = dataRow.createCell(cell++);
                        try
                        {
                            Object cellValue = null;
                            if(data instanceof Map){
                                cellValue = ((Map)data).get(field.getName());
                            }else if(data instanceof IExcelEntity){
//                                cellValue = BeanUtils.getProperty(data, field.getName());
                                cellValue = PropertyUtils.getProperty(data, field.getName());
                            }else {
                                throw new ExcelGuideException("不支持的实体类型:" + data.getClass().getName() + "必须实现接口" + IExcelEntity.class.getName());
                            }
                            //如果值为空，且默认值不为空，则取默认值替代
                            if(StringUtils.isBlank(String.valueOf(cellValue)) && StringUtils.isNotBlank(field.getDefaultValue())){
                            	cellValue = field.getDefaultValue();
                            }
                            if(cellValue != null){
                            	//列宽不为空
                            	if(field.getWidthOfColumn() != null && field.getWidthOfColumn() != 0){
                            		sheet.setColumnWidth(cell - 1, 255 * field.getWidthOfColumn()); //设置该列宽度,以一个字符的宽度为单位
                            	}
                            	
                                //如果是需要转换的类型，则取转换后的值,引时不再判断数据类型，
                                if(field.getConvert()){
                                    cellValue = field.getEntrys().get(cellValue.toString());
                                    Object[] object = field.getEntrys().values().toArray();
                                    String[] strs = Arrays.asList(object).toArray(new String[0]);
                                    sheet.addValidationData(ExcelBuilder.setDataValidation(sheet,strs,1,5000,field.getIndex()-1 ,field.getIndex()-1));
                                    if(cellValue != null){
                                    	  dataCell.setCellValue(ExcelBuilder.getRichTextString(workbook, String.valueOf(cellValue)));
                                    }else{
                                    	  dataCell.setCellValue(ExcelBuilder.getRichTextString(workbook, ""));
                                    }
                                  
                                }else{
                                  //根据数据类型填充表格数据
                                    if(ExcelConstants.DATA_TYPE_DATE.equalsIgnoreCase(field.getDataType())){
                                        String dateStr = DateHelper.format((Date)cellValue, field.getFormat());
                                        dataCell.setCellValue(ExcelBuilder.getRichTextString(workbook, dateStr));
                                    }
                                    else if(ExcelConstants.DATA_TYPE_INTEGER.equalsIgnoreCase(field.getDataType()))
                                    {
                                        dataCell.setCellValue(NumberUtils.toInt(cellValue.toString(), 0));
                                    }else if(ExcelConstants.DATA_TYPE_DOUBLE.equalsIgnoreCase(field.getDataType()))
                                    {
                                        dataCell.setCellValue(NumberUtils.toDouble(cellValue.toString(),0d));
                                    }else if(ExcelConstants.DATA_TYPE_LONG.equalsIgnoreCase(field.getDataType()))
                                    {
                                        dataCell.setCellValue(NumberUtils.toLong(cellValue.toString(),0l));
                                    }
                                    else if(ExcelConstants.DATA_TYPE_IMAGE.equalsIgnoreCase(field.getDataType())){
                                    	byte[] imageByte = null;
                                    	if(cellValue instanceof File){
                                    		File image = (File)cellValue;
                                    		if(image != null && image.length() > 0)
                                    			imageByte = IOUtils.toByteArray(new FileInputStream((File)cellValue));
//                                    			bufferImg = ImageIO.read(image);
                                    	}else if(cellValue instanceof BufferedImage){
                                    		BufferedImage bufferImg = (BufferedImage)cellValue;
                                    		//图片的导出
                                        	ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
                                        	ImageIO.write(bufferImg, field.getImageType().name(), byteArrayOut);
                                        	imageByte = byteArrayOut.toByteArray();
                                    	}else{
                                    		throw new ExcelGuideException("不支持的图片数据类型,允许为File/BufferedImage");
                                    	}
                                    	
                                    	/**
                                    	 * 关于HSSFClientAnchor(dx1,dy1,dx2,dy2,col1,row1,col2,row2)的参数：
											dx1：起始单元格的x偏移量，如例子中的255表示直线起始位置距A1单元格左侧的距离；
											dy1：起始单元格的y偏移量，如例子中的125表示直线起始位置距A1单元格上侧的距离；
											dx2：终止单元格的x偏移量，如例子中的1023表示直线起始位置距C3单元格左侧的距离；
											dy2：终止单元格的y偏移量，如例子中的150表示直线起始位置距C3单元格上侧的距离；
											col1：起始单元格列序号，从0开始计算；
											row1：起始单元格行序号，从0开始计算，如例子中col1=0,row1=0就表示起始单元格为A1；
											col2：终止单元格列序号，从0开始计算；
											row2：终止单元格行序号，从0开始计算，如例子中col2=2,row2=2就表示起始单元格为C3；
                                    	 * 
                                    	 */
                                    	ClientAnchor anchor = null;//new HSSFClientAnchor(0 , 0, 1023 , 253,  col1,  row1, col2, row2);//dx2最大值 1023,dy2最大值255
                                    	int PICTURE_TYPE = XSSFWorkbook.PICTURE_TYPE_PNG;
                                    	if(patriarch instanceof XSSFDrawing){
                                    		anchor = new XSSFClientAnchor(0, 0, 0, 0,  (short)(cell - 1),  rowIndex - 1, cell, rowIndex);
                                    	} else{
                                    		anchor = new HSSFClientAnchor(0, 0, 0, 0,  (short)(cell - 1),  rowIndex - 1, cell, rowIndex);//dx2最大值 1023,dy2最大值255
                                    	}
                                    	//设置图片类型
                                    	if(field.getImageType() == ImageType.JPG){
                                    		PICTURE_TYPE = XSSFWorkbook.PICTURE_TYPE_JPEG;
                                    	}
                                    	patriarch.createPicture(anchor,workbook.addPicture(imageByte, PICTURE_TYPE));
                                    }
                                    else{
                                    		 dataCell.setCellValue(ExcelBuilder.getRichTextString(workbook, cellValue.toString()));
                                    }
                                }
                            }
                            else{
                                dataCell.setCellValue(ExcelBuilder.getRichTextString(workbook, ""));
                            }
                            
                        }
                        catch (Exception e)
                        {
                            dataCell.setCellStyle(ExcelBuilder.buildErrorStyle(workbook));
//                            dataCell.setCellValue(new HSSFRichTextString("[ERROR]" + e.getMessage()));
                            dataCell.setCellValue("[ERROR]" + e.getMessage());
                            
                            if(patriarch == null)
                            	patriarch = sheet.createDrawingPatriarch();
                         // 定义注释的大小和位置,详见文档
//                            HSSFComment comment = patriarch.createComment(new HSSFClientAnchor(rowIndex ,cell -1 , rowIndex , cell -1 , (short) 3, 2, (short) 6, 5));
                            Comment comment = patriarch.createCellComment(patriarch.createAnchor(rowIndex ,cell -1 , rowIndex , cell -1 , (short) 3, 2, (short) 6, 5));
                         // 设置注释内容
                            comment.setString(ExcelBuilder.getRichTextString(workbook, "导出表格数据出错：" + e.getMessage()));
                         // 设置注释作者，当鼠标移动到单元格上是可以在状态栏中看到该内容.
                            comment.setAuthor("Being");
//                            comment.setVisible(true);
                            dataCell.setCellComment(comment);
                            
                            logger.error("导出表格数据出错：" + field,e);
                        }
                    }
                }
            }
        }else{
            flag = false;
        }
        return flag;
    }

    /**
     * entity的getter方法
     * 
     * @return the entity
     */
    public GuideEntity getEntity()
    {
        return entity;
    }

    /**
     * entity的setter方法
     * 
     * @param entity the entity to set
     */
    public void setEntity(GuideEntity entity)
    {
        this.entity = entity;
    }

    /**
     * workbook的getter方法
     * 
     * @return the workbook
     */
    public Workbook getWorkbook()
    {
        return workbook;
    }

    /**
     * workbook的setter方法
     * 
     * @param workbook the workbook to set
     */
    public void setWorkbook(Workbook workbook)
    {
        this.workbook = workbook;
    }

}
