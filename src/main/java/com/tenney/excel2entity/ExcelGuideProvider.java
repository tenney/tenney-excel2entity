/**
 * 版权所有：tenney
 * 项目名称: excel-guide-plugin
 * 类名称:ExcelGuideProvider.java
 * 包名称:com.tenney.excel2entity
 * 
 * 创建日期:2013年10月17日 下午8:02:07
 * 创建人： 唐雄飞		
 * <修改人>      <时间>      <版本号>    <描述>
 * 唐雄飞      2013年10月17日     	V1.0.0        N/A
 */

package com.tenney.excel2entity;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang3.BooleanUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.dom4j.Document;
import org.dom4j.Element;
import org.springframework.core.io.Resource;

import com.tenney.excel2entity.ExcelConstants.ImageType;
import com.tenney.excel2entity.lang.ExcelGuideException;
import com.tenney.excel2entity.lang.excel.ExcelBuilder;
import com.tenney.excel2entity.lang.xml.XmlParser;
import com.tenney.excel2entity.support.GuideEntity;
import com.tenney.excel2entity.support.GuideEntityField;
import com.tenney.excel2entity.support.IExcelReadCallBack;


/**
 * 类说明: <br/>
 * 创建人: 唐雄飞 <br/>
 * 创建日期:2013年10月17日 <br/> 
 * 
 */
public class ExcelGuideProvider
{
    private Logger logger = Logger.getLogger(this.getClass());
    /**
     * 配置文件
     */
    private String guideConfig;
    private File guideConfigFile;
    private Resource[] locations;
    
    public ExcelGuideProvider(String guideConfig) throws Exception{
        this.guideConfig = guideConfig;
        this._init();
    }
    
    public ExcelGuideProvider(File guideConfigFile) throws Exception{
        this.guideConfigFile = guideConfigFile;
        this._init();
    }
    
    public ExcelGuideProvider(Resource[] locations) throws Exception{
        this.locations = locations;
        this._init();
    }
    
    /**
     * 
     * 方法描述:获取第一个sheet的最大行号，最后一行的行号，该方法不处理有格式的空格<br>
     * 创建人:唐雄飞<br>
     * 创建日期:2015年11月29日<br>
     * @return<br>
     * @throws IOException 
     * @throws InvalidFormatException 
     */
    public Integer getMaxRows(InputStream is, boolean closeOnFinshed) throws InvalidFormatException, IOException{
    	try{
    		return new ExcelImportFromInput(ExcelBuilder.readWorkbook(is)).getMaxRows();
    	}finally{
    		if(closeOnFinshed)
    			IOUtils.closeQuietly(is);
    	}
    }
    /**
     * 
     * 方法描述:获取第一个sheet的最大行号，最后一行的行号，该方法不处理有格式的空格<br>
     * 创建人:唐雄飞<br>
     * 创建日期:2015年11月29日<br>
     * @param excelFile
     * @return
     * @throws InvalidFormatException
     * @throws IOException<br>
     */
    public Integer getMaxRows(File excelFile) throws InvalidFormatException, IOException{
    	
    	FileInputStream fis = null;
    	try{
    		fis = new FileInputStream(excelFile);
    		return new ExcelImportFromInput(ExcelBuilder.readWorkbook(fis)).getMaxRows();
    	}finally{
    		IOUtils.closeQuietly(fis);
    	}
    }
    
    /**
     * 
     * 方法描述: 从指定文件中按配置读取数据 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月24日 <br>
     * @param is  文件流对象 
     * @param guideId 配置标识
     * @return 数据集合<br>
     * @throws ExcelGuideException 
     */
    public <T> Collection<T> readFromExcel(InputStream is ,String guideId) throws ExcelGuideException{
        return this.readFromExcel(is, guideId, null,null);
    }
    
    /**
     * 
     * 方法描述: 从指定文件中按配置读取数据 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月24日 <br>
     * @param is 文件流
     * @param guideId 配置项标识
     * @param outFile 处理结果存放流
     * @return 成功的结果集合 <br>
     * @throws ExcelGuideException 
     */
    public <T> Collection<T> readFromExcel(InputStream is ,String guideId,OutputStream outFile,IExcelReadCallBack<T> callBack) throws ExcelGuideException{
        if(StringUtils.isNotBlank(guideId)){
            GuideEntity entity = ExcelConstants.guides.get(guideId);
            if(entity == null){
                throw new ExcelGuideException("错误的配置项ID，不存在该配置，请检查参数!");
            }
            try
            {
                Workbook workbook = ExcelBuilder.readWorkbook(is);
                Collection<T> dataSet = this._readExcel(entity,workbook,callBack);
                if(outFile != null && workbook != null){
                    workbook.write(outFile);
                }
                return dataSet;
            } catch (Exception e)
            {
                throw new ExcelGuideException("读取Excel文件出错:" + e.getMessage(),e);
            }
        }else{
            throw new ExcelGuideException("参数空，必须指定配置项ID!");
        }
    }
    
    /**
     * 
     * 方法描述: 从指定文件中按配置读取数据 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月24日 <br>
     * @param excelFile 文件信息
     * @param guideId 配置标识
     * @return 数据集合
     * @throws FileNotFoundException <br>
     * @throws ExcelGuideException 
     */
    public <T> Collection<T> readFromExcel(File excelFile, String guideId) throws FileNotFoundException, ExcelGuideException
    {
        return this.readFromExcel(new FileInputStream(excelFile), guideId);
    }
    /**
     * 
     * 方法描述: 从指定文件中按配置读取数据  <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月24日 <br>
     * @param excelFile 文件信息
     * @param guideId 配置标识
     * @param outFile 处理结果存放流
     * @return 数据集合
     * @throws FileNotFoundException <br>
     * @throws ExcelGuideException 
     */
    public <T> Collection<T> readFromExcel(File excelFile, String guideId, OutputStream outFile)
            throws FileNotFoundException, ExcelGuideException
    {
        return this.readFromExcel(excelFile, guideId, outFile, null);
    }
    
    /**
     * 
     * 方法描述: 从指定文件中按配置读取数据 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月24日 <br>
     * @param excelFile 文件路径
     * @param guideId 配置标识
     * @return 数据集合
     * @throws FileNotFoundException <br>
     * @throws ExcelGuideException 
     */
    public <T> Collection<T> readFromExcel(String excelFile,String guideId) throws FileNotFoundException, ExcelGuideException{
        return this.readFromExcel(new FileInputStream(excelFile), guideId);
    }
    
    public <T> Collection<T> readFromExcel(File excelFile, String guideId, OutputStream outFile,IExcelReadCallBack<T> callBack)
            throws FileNotFoundException, ExcelGuideException
    {
        return this.readFromExcel(new FileInputStream(excelFile), guideId, outFile,callBack);
    }
    
    /**
     * 
     * 方法描述: <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月24日 <br>
     * @param entity
     * @param workbook
     * @return <br>
     * @throws ExcelGuideException 
     */
    private <T> Collection<T> _readExcel(GuideEntity entity , Workbook workbook,IExcelReadCallBack<T> callBack) throws ExcelGuideException{
        try
        {
            return new ExcelImportFromInput(entity,workbook).invokeImport(callBack);
        }  catch (Exception e)
        {
            throw new ExcelGuideException(e);
        }
    }
    
    
    /**
     * 
     * 方法描述: 根据请求响应对象输出EXCEL文档 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月23日 <br>
     * @param response
     * @param guideId
     * @param dataSet
     * @throws ExcelGuideException
     * @throws IOException
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     * @throws NoSuchMethodException <br>
     */
    public void WriteToExcel(HttpServletResponse response,String guideId, Collection<?> dataSet,String fileSuffix) throws ExcelGuideException, IOException, IllegalAccessException, InvocationTargetException, NoSuchMethodException{
        if(StringUtils.isNotBlank(guideId)){
            GuideEntity entity = ExcelConstants.guides.get(guideId);
            if(entity == null){
                throw new ExcelGuideException("错误的配置项ID，不存在该配置，请检查参数!");
            }
            Workbook workbook = this._writeData(entity, dataSet,fileSuffix);
            
            //根据workbook类型产生Excel文件后缀名(xls,xlsx)
            String EXCEL_FILE_SUFFIX = "." + ExcelConstants.EXCEL_FILE_SUFFIX_XLSX;
            if(workbook instanceof HSSFWorkbook){
            	EXCEL_FILE_SUFFIX = "." + ExcelConstants.EXCEL_FILE_SUFFIX_XLS;
            }
            String fileName = (StringUtils.isNotBlank(entity.geteName())?entity.geteName():entity.getEid()) + EXCEL_FILE_SUFFIX;
            logger.info("导入Excel:" + fileName);
            response.addHeader("Content-Disposition", "attachment; filename=\"" + new String(fileName.getBytes(), "ISO-8859-1") + "\"");
            response.setContentType("octets/stream");
//            response.setContentType("application/x-msexcel");
//            response.addHeader("Content-Length", String.valueOf(downloadFile.length()));
//            this.WriteToExcel(response.getOutputStream(), guideId, dataSet);
            
            workbook.write(response.getOutputStream());
            response.getOutputStream().flush();//刷新输出流
            response.getOutputStream().close();//关闭输出流
        }else{
            throw new ExcelGuideException("参数空，必须指定配置项ID!");
        }
    }
    
    /**
     * 
     * 方法描述: 将数据写入到表格 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月25日 <br>
     * @param os 输出流，表格最后输出
     * @param guideId 配置ID
     * @param dataSet 数据集
     * @param fileSuffix 文件后缀名(xls,xlsx)
     * @throws ExcelGuideException
     * @throws IOException
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     * @throws NoSuchMethodException <br>
     */
    public void WriteToExcel(OutputStream os,String guideId, Collection<?> dataSet,String fileSuffix) throws ExcelGuideException, IOException, IllegalAccessException, InvocationTargetException, NoSuchMethodException{
        if(StringUtils.isNotBlank(guideId)){
            GuideEntity entity = ExcelConstants.guides.get(guideId);
            if(entity == null){
                throw new ExcelGuideException("错误的配置项ID，不存在该配置，请检查参数!");
            }
            Workbook workbook = this._writeData(entity, dataSet,fileSuffix);
            workbook.write(os);
            os.flush();
            os.close();
        }else{
            throw new ExcelGuideException("参数空，必须指定配置项ID!");
        }
    }
    
    /**
     * 
     * 方法描述: <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月25日 <br>
     * @param guideId
     * @param dataSet
     * @return
     * @throws ExcelGuideException
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     * @throws NoSuchMethodException <br>
     */
    public Workbook WriteToExcel(String guideId, Collection<?> dataSet,String fileSuffix) throws ExcelGuideException, IllegalAccessException, InvocationTargetException, NoSuchMethodException{
        if(StringUtils.isNotBlank(guideId)){
            GuideEntity entity = ExcelConstants.guides.get(guideId);
            if(entity == null){
                throw new ExcelGuideException("错误的配置项ID，不存在该配置，请检查参数!");
            }
            return this._writeData(entity, dataSet,fileSuffix);
        }else{
            throw new ExcelGuideException("参数空，必须指定配置项ID!");
        }
    }
    
    private Workbook _writeData(GuideEntity entity,Collection<?> dataSet,String fileSuffix) throws ExcelGuideException, IllegalAccessException, InvocationTargetException, NoSuchMethodException{
        Workbook workbook = ExcelBuilder.getWorkBookInstance(fileSuffix);
        if(new ExcelExportToExcel(entity,workbook).invokeExport(dataSet)){
            return workbook;
        }else{
            throw new ExcelGuideException("导出数据出错!");
        }
    }
    
    /**
     * 
     * 方法描述: 初始化配置 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月18日 <br> <br>
     * @throws Exception 
     */
    private void _init() throws Exception{
        logger.debug("Loading Excel-guide-plugin ...");
        List<Document> configs = new ArrayList<Document>();
        //解析locations
        if(this.locations != null && this.locations.length > 0){
            for(Resource source:this.locations){
                configs.add(XmlParser.parseFromSource(source.getInputStream()));
                logger.info("Loading " + source.getFilename());
            }
        }
        //解析 guideConfig
        if(StringUtils.isNotBlank(this.guideConfig)){
            configs.add(XmlParser.parseFromSource(this.guideConfig));
            logger.info("Loading " + this.guideConfig);
        }
        //解析 guideConfigFile
        if(this.guideConfigFile != null){
            configs.add(XmlParser.parseFromSource(new FileInputStream(this.guideConfigFile)));
            logger.info("Loading " + this.guideConfigFile.getName());
        }
        
        this.buildGuides(configs);
    }
    
    /**
     * 
     * 方法描述: 解析配置文件并生成配置 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月18日 <br> <br>
     * @throws Exception 
     */
    @SuppressWarnings("unchecked")
    private void buildGuides(List<Document> docs) throws Exception{
        if(docs != null && !docs.isEmpty()){
            for(Document document: docs){
                try
                {
                    Element root = document.getRootElement();
                    if(!ExcelConstants.GUIDE_CONFIG_ROOT.equals(root.getName())){
                        logger.warn("不合规则的配置文件：" + root.getName());
                        continue;
                    }
                    Iterator<Element> itElements = root.elementIterator(ExcelConstants.GUIDE_CONFIG_ELEMENT);
                    while (itElements.hasNext()){
                        //解析实体配置
                        Element eEntity = itElements.next();
                        String eid = StringUtils.trim(eEntity.attributeValue(ExcelConstants.GUIDE_CONFIG_ELEMENT_ID));
                        
                        if(StringUtils.isBlank(eid)){
                            throw new ExcelGuideException("实体标识id必须指定:" + eEntity.getName());
                        }
                        if(ExcelConstants.guides.get(eid) != null){
                            throw new ExcelGuideException("存在相同标识的实体配置:" + eid);
                        }
                        GuideEntity entity = new GuideEntity();
                        entity.setEid(eid);
                        
                        String eName = StringUtils.trim(eEntity.attributeValue(ExcelConstants.GUIDE_CONFIG_ELEMENT_NAME));
                        entity.seteName(eName);
                        
                        logger.debug("加载Excel配置：[" + eName + " --> " + eid + "]");
                        
                        String entityClass = StringUtils.trim(eEntity.attributeValue(ExcelConstants.GUIDE_CONFIG_ELEMENT_CLASS));
                        if(StringUtils.isBlank(entityClass)){
                            throw new ExcelGuideException("实体类型class必须指定:" + eEntity.getName());
                        }
                        entity.setEntityClass(entityClass);
                        //设置每行的高度
                        entity.setHeightOfRows(NumberUtils.createFloat(eEntity.attributeValue(ExcelConstants.GUIDE_CONFIG_ELEMENT_ROW_HEIGHT)));
                        
                        //解析实体字段
                        Iterator<Element> itFields = eEntity.elementIterator(ExcelConstants.GUIDE_CONFIG_FIELD);
                        while(itFields.hasNext()){
                            Element eField = itFields.next();
                            String fName = StringUtils.trim(eField.attributeValue(ExcelConstants.GUIDE_CONFIG_FIELD_NAME));
                            if(StringUtils.isBlank(fName)){
                                throw new ExcelGuideException("必须指定字段名: " + eid + " -> " + eField.getName() + " -> " + ExcelConstants.GUIDE_CONFIG_FIELD_NAME);
                            }
                            GuideEntityField field = new GuideEntityField();
                            //字段名
                            field.setName(fName);
                            //表格标题
                            field.setExcelTitle(StringUtils.trim(eField.attributeValue(ExcelConstants.GUIDE_CONFIG_FIELD_EXCEL)));
                            //排序
                            field.setIndex(NumberUtils.createInteger(StringUtils.trim(eField.attributeValue(ExcelConstants.GUIDE_CONFIG_FIELD_INDEX))));
                            //列宽
                            field.setWidthOfColumn(NumberUtils.toShort(eField.attributeValue(ExcelConstants.GUIDE_CONFIG_FIELD_WIDTH)));
                            //默认值 
                            field.setDefaultValue(StringUtils.trim(eField.attributeValue(ExcelConstants.GUIDE_CONFIG_FIELD_DFVALUE)));
                            //数据类型
                            field.setDataType(StringUtils.trim(eField.attributeValue(ExcelConstants.GUIDE_CONFIG_FIELD_DATATYPE)));
                            if(ExcelConstants.DATA_TYPE_IMAGE.equals(field.getDataType())){
                            	//图片类型
                                String imageType = StringUtils.trimToEmpty(eField.attributeValue(ExcelConstants.GUIDE_CONFIG_FIELD_IMAGETYPE));
                                if(StringUtils.isBlank(imageType)){
                                	throw new Exception("图片字段必须设置图片类型(imageType)字段!");
                                }
                                try {
    								field.setImageType(ImageType.valueOf(imageType.toUpperCase()));
    							} catch (Exception e) {
    								throw new Exception("不支持的图片类型，可选类型为jpg/png");
    							}
                                
                                field.setImported(false); //图片字段提供导入
                            }
                            
                            field.setNullable(BooleanUtils.toBoolean(StringUtils.trim(eField.attributeValue(ExcelConstants.GUIDE_CONFIG_FIELD_NULLABLE))));
                            //是否允许为空
                            if(!field.getNullable() && StringUtils.isBlank(field.getDefaultValue())){
                                logger.debug("非空字段未指定默认值:" + eid + " -> " + fName + " -> " + ExcelConstants.GUIDE_CONFIG_FIELD_NULLABLE);
//                                throw new ExcelGuideException("非空字段必须指定一个默认值:" + eid + " -> " + fName + " -> " + ExcelConstants.GUIDE_CONFIG_FIELD_NULLABLE);
                            }
                            //格式
                            field.setFormat(StringUtils.trim(eField.attributeValue(ExcelConstants.GUIDE_CONFIG_FIELD_FORMAT)));
                            if(ExcelConstants.DATA_TYPE_DATE.equalsIgnoreCase(field.getDataType()) && StringUtils.isBlank(field.getFormat())){
                                throw new ExcelGuideException("日期类型字段必须指定格式:" + eid + " -> " + fName + " -> " + ExcelConstants.GUIDE_CONFIG_FIELD_FORMAT);
                            }
                            
                            //导入是否需要提供该字段
                            String imported = StringUtils.trim(eField.attributeValue(ExcelConstants.GUIDE_CONFIG_FIELD_IMPORTED));
                            field.setImported(BooleanUtils.toBoolean(imported));
                            
                            //是否需要转换
                            String convert = StringUtils.trim(eField.attributeValue(ExcelConstants.GUIDE_CONFIG_FIELD_CONVERT));
                            if(StringUtils.isNotBlank(convert)){
                                field.setConvert(BooleanUtils.toBoolean(convert));
                            }else{
                                field.setConvert(false);
                            }
                            //字段是否需要转换
                            if(field.getConvert()){
                                Iterator<Element> entrys = eField.elementIterator(ExcelConstants.GUIDE_CONFIG_FIELD_ENTRY);
                                while(entrys.hasNext()){
                                    Element keyMap = entrys.next();
                                    String key = StringUtils.trim(keyMap.attributeValue(ExcelConstants.GUIDE_CONFIG_ENTRY_KEY));
                                    String value = StringUtils.trim(keyMap.attributeValue(ExcelConstants.GUIDE_CONFIG_ENTRY_VALUE));
                                    if(StringUtils.isNotBlank(key) && StringUtils.isNotBlank(value)){
                                        field.getEntrys().put(key, value);
                                    }else{
                                        throw new ExcelGuideException("属性值必须全部指定:" + eid + " -> " + fName + " -> " + ExcelConstants.GUIDE_CONFIG_FIELD_ENTRY);
                                    }
                                }
                                if(field.getEntrys().isEmpty()){
                                    throw new ExcelGuideException("需要转换的字段必须指定键值映射:" + eid + " -> " + fName + " -> " + ExcelConstants.GUIDE_CONFIG_FIELD_ENTRY);
                                }
                            }
                            
                            entity.getFields().add(field);
                        }
                        
                        if(entity.getFields().isEmpty()){
                            throw new ExcelGuideException("必须提供至少一个字段的配置:" + eid);
                        }
                        //将配置加入集合保存
                        ExcelConstants.guides.put(eid, entity);
                    }
                } catch (Exception e)
                {
                    logger.error("解析配置文件出错：" + e.getMessage() ,e);
                    throw new Exception("解析配置文件出错：" + e.getMessage() ,e);
                }
            }
        }else{
            logger.warn("未加载任何配置信息 .");
        }
    }
    
    /**
     * guideConfig的getter方法
     * @return the guideConfig
     */
    public String getGuideConfig()
    {
        return guideConfig;
    }

    /**
     * guideConfigFile的getter方法
     * @return the guideConfigFile
     */
    public File getGuideConfigFile()
    {
        return guideConfigFile;
    }

    /**
     * locations的getter方法
     * @return the locations
     */
    public Resource[] getLocations()
    {
        return locations;
    }
}
