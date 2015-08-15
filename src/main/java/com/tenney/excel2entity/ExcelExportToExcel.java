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

import java.lang.reflect.InvocationTargetException;
import java.util.Collection;
import java.util.Date;
import java.util.Map;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.tenney.excel2entity.lang.DateHelper;
import com.tenney.excel2entity.lang.ExcelGuideException;
import com.tenney.excel2entity.lang.excel.ExcelBuilder;
import com.tenney.excel2entity.support.GuideEntity;
import com.tenney.excel2entity.support.GuideEntityField;
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

            //填充标题 
            Row titleRow = sheet.createRow(ExcelConstants.TITLE_ROW_IDX);
            short idx = 0;
            for(GuideEntityField field : this.entity.getFields()){
                Cell cell = titleRow.createCell(idx++);
                cell.setCellStyle(titleStyle);
                if(!field.getImported()){
//                	cell.setCellComment();
                }
//                HSSFRichTextString fieldText = new HSSFRichTextString(field.getExcelTitle());
//                cell.setCellValue(fieldText);
                cell.setCellValue(field.getExcelTitle());
            }
            
//            ConvertUtils.register(converter, clazz);
            
            //填充数据内容
            if(dataSet != null && !dataSet.isEmpty()){
                int rowIndex = ExcelConstants.TITLE_ROW_IDX + 1;
                for(Object data:dataSet){
                    short cell = 0;
                    //创建行，在标题行下面
                    Row dataRow = sheet.createRow(rowIndex++);
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
                                //如果是需要转换的类型，则取转换后的值,引时不再判断数据类型，
                                if(field.getConvert()){
                                    cellValue = field.getEntrys().get(cellValue.toString());
                                    dataCell.setCellValue(ExcelBuilder.getRichTextString(workbook, String.valueOf(cellValue)));
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
                                    else{
                                        dataCell.setCellValue(ExcelBuilder.getRichTextString(workbook, cellValue.toString()));
                                    }
                                }
                            }
                            else{
                                dataCell.setCellValue(ExcelBuilder.getRichTextString(workbook, null));
                            }
                            
                        }catch(ExcelGuideException e){
                            //对象不支持的错，继续往外抛出，避免错误数据继续读取
                            throw new ExcelGuideException(e);
                        } catch (Exception e)
                        {
                            dataCell.setCellStyle(ExcelBuilder.buildErrorStyle(workbook));
//                            dataCell.setCellValue(new HSSFRichTextString("[ERROR]" + e.getMessage()));
                            dataCell.setCellValue("[ERROR]" + e.getMessage());
                            
                            Drawing patriarch = sheet.createDrawingPatriarch();
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
                            e.printStackTrace();
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
