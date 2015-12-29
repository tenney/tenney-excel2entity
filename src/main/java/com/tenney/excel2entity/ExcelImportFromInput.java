/**
 * 版权所有：tenney
 * 项目名称: eicsp
 * 类名称:ExcelImportFromInput.java
 * 包名称:com.tenney.excel2entity
 * 
 * 创建日期:2013年10月24日 下午5:03:53
 * 创建人： 唐雄飞		
 * <修改人>      <时间>      <版本号>    <描述>
 * 唐雄飞      2013年10月24日     	V1.0.0        N/A
 */

package com.tenney.excel2entity;

import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.tenney.excel2entity.lang.DateHelper;
import com.tenney.excel2entity.lang.ExcelGuideException;
import com.tenney.excel2entity.lang.excel.ExcelBuilder;
import com.tenney.excel2entity.support.GuideEntity;
import com.tenney.excel2entity.support.GuideEntityField;
import com.tenney.excel2entity.support.ICallBackMessage;
import com.tenney.excel2entity.support.IExcelEntity;
import com.tenney.excel2entity.support.IExcelReadCallBack;
import com.tenney.excel2entity.support.IExcelReadInterruptCallBack;

/**
 * 类说明: <br/>
 * 创建人: 唐雄飞 <br/>
 * 创建日期:2013年10月24日 <br/> 
 * 
 */
public class ExcelImportFromInput
{
    private static final Logger logger = Logger.getLogger(ExcelImportFromInput.class);
    
    private GuideEntity entity;
    private Workbook workbook;
    
    /**
     * ExcelImportFromInput.java的构造函数
     */
    public ExcelImportFromInput()
    {
        super();
    }
    
    /**
     * 
     * ExcelImportFromInput.java的构造函数
     * @param workbook
     */
    public ExcelImportFromInput(Workbook workbook){
    	super();
        this.workbook = workbook;
    }
    
    /**
     * ExcelImportFromInput.java的构造函数
     * @param entity
     * @param workbook
     */
    public ExcelImportFromInput(GuideEntity entity, Workbook workbook)
    {
        super();
        this.entity = entity;
        this.workbook = workbook;
    }
    
    /**
     * 
     * 方法描述: 执行Excel表格内容数据解析，返回数据集 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月24日 <br>
     * @return <br>
     * @throws ExcelGuideException 
     * @throws InstantiationException 
     * @throws InvocationTargetException 
     * @throws IllegalAccessException 
     * @throws NoSuchMethodException 
     * @throws ClassNotFoundException 
     */
    @SuppressWarnings({ "unchecked", "rawtypes" })
    public <T> Collection<T> invokeImport(IExcelReadCallBack<T> callBack) throws ExcelGuideException, NoSuchMethodException, IllegalAccessException, InvocationTargetException, InstantiationException, ClassNotFoundException{
        List<T> dataSet = new ArrayList<T>();
        if(null != this.workbook && null != this.entity){
            //获取第一个sheet
            Sheet sheet = this.workbook.getSheetAt(0);
            Row titleRow = sheet.getRow(ExcelConstants.TITLE_ROW_IDX);
            int colNum = titleRow.getPhysicalNumberOfCells();
            if(colNum != this.entity.getFields().size() || 0 == colNum){
                throw new ExcelGuideException("表格内容列数不符合配置的表格要求!");
            }
            //标题名与列索引对应的Map集合
            Map<Short,String> titleMap = new HashMap<Short,String>();
            for(Short i = 0; i <= colNum; i++){
                Cell title = titleRow.getCell(i);
                if(title == null) continue;
                titleMap.put(i, title.getRichStringCellValue().getString());
            }
            //获取最后一行的行号
            int lastRowIndex = sheet.getLastRowNum();
            //遍历表格
            for(int idx = ExcelConstants.TITLE_ROW_IDX + 1 ; idx <= lastRowIndex ; idx++){
                titleRow = sheet.getRow(idx);
                String titleName = "";
                Short cols = 0;
                try
                {
                    if(titleRow == null || titleRow.getFirstCellNum() < 0) continue;  //如果空行，直接跳过
                    Object voBean = null;
                    //兼容Map接口配置
                    String clazz = this.entity.getEntityClass();
                    if(clazz.equalsIgnoreCase("java.util.HashMap") || clazz.equalsIgnoreCase("java.util.Map") || clazz.equalsIgnoreCase("Map") || clazz.equalsIgnoreCase("HashMap")){
                    	voBean = new HashMap();
                    }else{
                    	voBean = Class.forName(clazz).newInstance();
                    }
                    for(Short r = 0; r < colNum; r++){
                    	cols = r;
                        Cell content = titleRow.getCell(r);
                        
                        //根据标题字符串，取对应的配置项字段
                        titleName = titleMap.get(r);
                        GuideEntityField cField = this.entity.getFieldMap().get(titleName.trim());
                        if(cField == null){
                            throw new ExcelGuideException("无效的字段:[行:" + idx + ",列：" + r + "]-->" + titleName);
                        }
                        Object readData = null;
                        //如果需要反转值
                        if(cField.getConvert()){
//                            String keyValue = content.getRichStringCellValue().getString();
                            String keyValue = String.valueOf(ExcelBuilder.getCellValue(content));
                            readData = cField.getEntryKey(keyValue);
                        }else{
                            if(ExcelConstants.DATA_TYPE_DATE.equals(cField.getDataType())){
//                                RichTextString rString = content.getRichStringCellValue();
                                Object cellValue = ExcelBuilder.getCellValue(content);
                                if(cellValue instanceof Date){
                                    readData = cellValue;
                                }else{
                                    readData = DateHelper.parse(String.valueOf(cellValue), cField.getFormat());
                                }
                            }
//                            else if(ExcelConstants.DATA_TYPE_INTEGER.equals(cField.getDataType())){
//                                HSSFRichTextString rString = content.getRichStringCellValue();
//                                readData = Integer.valueOf(rString.getString());
//                            }
                            else if(ExcelConstants.DATA_TYPE_DOUBLE.equals(cField.getDataType())){
                                readData = content.getNumericCellValue();
                            }
//                            else if(ExcelConstants.DATA_TYPE_LONG.equalsIgnoreCase(cField.getDataType())){
//                                HSSFRichTextString rString = content.getRichStringCellValue();
//                                readData = Long.valueOf(rString.getString());
//                            }
                            else{
//                                RichTextString rString = content.getRichStringCellValue();
//                                readData = rString.getString();
                                readData = String.valueOf(ExcelBuilder.getCellValue(content));
                            }
                        }
                        //如果为空，则取默认值
                        if(readData == null){
                            //如果是日期类型，且默认值为 "sysDate"，则直接格式化日期
                            if(ExcelConstants.DATA_TYPE_DATE.equalsIgnoreCase(cField.getDataType()) && ExcelConstants.DATA_TYPE_DATE_SYSDATE.equalsIgnoreCase(cField.getDefaultValue())){
                                readData = new Date();
                            }else{
                                readData = cField.getDefaultValue();
                            }
                        }
                        //判断非空值字段
                        if(readData == null && !cField.getNullable()){
                            throw new ExcelGuideException("该字段为必填:" + cField.getExcelTitle());
                        }
                        /**
                         * 属性赋值
                         */
                        if(voBean instanceof IExcelEntity){
//                            PropertyUtils.setProperty(voBean, cField.getName(), readData);
                            BeanUtils.setProperty(voBean, cField.getName(), readData);
                        }else{
                            Map<String,Object> bean = (Map<String,Object>)voBean;
                            bean.put(cField.getName(), readData);
                        }
                    }
                    //添加数据
                    dataSet.add((T)voBean);
                    //方法回调 
                    if(callBack != null){
                        try{
                        	//回调方法
                            ICallBackMessage message = callBack.mapRow((T)voBean,workbook);
                            if(message != null){
                                Cell messageCell = titleRow.createCell((short)(colNum));
                                messageCell.setCellValue(ExcelBuilder.getRichTextString(workbook, message.getMessage()));
                                if(message.getStyle() != null){
                                    messageCell.setCellStyle(message.getStyle());
                                }
                            }
                        }catch(Exception e){
                        	if(callBack instanceof IExcelReadInterruptCallBack)
                        	{
                        		logger.error("记录读取被中断，原因：" + e.getMessage() ,e);
                        		throw new Exception("记录读取被中断，原因：" + e.getMessage());
                        	}
                        }
                    }
                } catch (Exception e)
                {
                    Cell errorCell = titleRow.createCell((short)(colNum));
                    errorCell.setCellStyle(ExcelBuilder.buildErrorStyle(workbook));
                    errorCell.setCellValue(ExcelBuilder.getRichTextString(workbook, "[失败]" + e.getMessage()));
                    logger.error("解析表格数据出错：" + e.getMessage(),e);
                    
                    if(callBack != null && (callBack instanceof IExcelReadInterruptCallBack)){
                    	throw new ExcelGuideException("解析表格数据出错：[行：" + idx + ",列：" + cols + "]-->" + titleName + "|" + e.getMessage());
                    }
                }
            }
        }
        return dataSet;
    }
    
    /**
     * 
     * 方法描述:获取第一个sheet的最大行号，最后一行的行号<br>
     * 创建人:唐雄飞<br>
     * 创建日期:2015年11月29日<br>
     * @return<br>
     */
    public Integer getMaxRows(){
    	if(null != this.workbook){
//    		Sheet sheet = this.workbook.getSheetAt(0);
//    		System.out.println(sheet.getPhysicalNumberOfRows());
//    		System.out.println(sheet.getLastRowNum());
//    		System.out.println("----------------------");
//    		try{
//    			int lastRowIndex = sheet.getLastRowNum();
//        		for(int idx = ExcelConstants.TITLE_ROW_IDX ; idx < lastRowIndex ; idx++){
//        			 
//        			Row r = sheet.getRow(idx); 
//        			if((r == null || r.getFirstCellNum() < 0) && idx < lastRowIndex ){
//        				logger.info("deleted row:" + (idx + 1));
//        				sheet.shiftRows(idx + 1, sheet.getLastRowNum(),-1);
//        			}
////        			System.out.println(r.getFirstCellNum());
//        			if(idx == lastRowIndex){
//        				Row removingRow=sheet.getRow(idx);
//            			if(removingRow != null)
//            				sheet.removeRow(removingRow);
//        			}
//        			
//        			lastRowIndex = sheet.getLastRowNum();
//        			
//        			if(r != null){
//        				System.out.println(r.getFirstCellNum() + "--------------");
//        				for(short i = 0 ; i < r.getLastCellNum(); i++){
//            				Cell c = r.getCell(i);
//            				System.out.print(c != null ? c.getStringCellValue():"" + "|");
//            			}
//            			System.out.println(" --");
//        			}
//        		}
//        		
//        		FileOutputStream fos = new FileOutputStream("C:\\Users\\tenney\\Desktop\\bb.xlsx");
//        		this.workbook.write(fos);
//        		fos.flush();
//        		fos.close();
//    		}catch(Exception e){
//    			logger.warn("清理表格空行出现异常！" + e.getMessage(), e);
//    		}
//    		return sheet.getLastRowNum();
    		return this.workbook.getSheetAt(0).getPhysicalNumberOfRows();
    	}
    	return 0;
    }
    
    /**
     * entity的getter方法
     * @return the entity
     */
    public GuideEntity getEntity()
    {
        return entity;
    }
    /**
     * entity的setter方法
     * @param entity the entity to set
     */
    public void setEntity(GuideEntity entity)
    {
        this.entity = entity;
    }
    /**
     * workbook的getter方法
     * @return the workbook
     */
    public Workbook getWorkbook()
    {
        return workbook;
    }
    /**
     * workbook的setter方法
     * @param workbook the workbook to set
     */
    public void setWorkbook(Workbook workbook)
    {
        this.workbook = workbook;
    }
    
    
}
