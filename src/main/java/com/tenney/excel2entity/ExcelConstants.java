/**
 * 版权所有：tenney
 * 项目名称: excel-guide-plugin
 * 类名称:ExcelConstants.java
 * 包名称:com.tenney.excel2entity
 * 
 * 创建日期:2013年10月17日 下午8:04:01
 * 创建人： 唐雄飞		
 * <修改人>      <时间>      <版本号>    <描述>
 * 唐雄飞      2013年10月17日     	V1.0.0        N/A
 */

package com.tenney.excel2entity;

import com.tenney.excel2entity.support.GuideEntity;

import java.util.HashMap;
import java.util.Map;

/**
 * 类说明: <br/>
 * 创建人: 唐雄飞 <br/>
 * 创建日期:2013年10月17日 <br/> 
 * 
 */
public class ExcelConstants
{
    private ExcelConstants(){}
    
    /**
     * 配置集合
     */
    public static final Map<String,GuideEntity> guides = new HashMap<String,GuideEntity>();
    
    /**
     * 文件后缀名
     */
    public static final String EXCEL_FILE_SUFFIX_XLS = "xls";
    public static final String EXCEL_FILE_SUFFIX_XLSX = "xlsx";
    
    /**
     * 表格标题所在行
     */
    public static final Integer TITLE_ROW_IDX = 0;
    
    /**
     * 配置文件根节点
     */
    public static final String GUIDE_CONFIG_ROOT = "guides";
    
    /**
     * 配置文件实体节点
     */
    public static final String GUIDE_CONFIG_ELEMENT = "entity";
    public static final String GUIDE_CONFIG_ELEMENT_ID = "id";
    public static final String GUIDE_CONFIG_ELEMENT_NAME = "name";
    public static final String GUIDE_CONFIG_ELEMENT_CLASS = "class";
    
    /**
     * 配置文件字段
     */
    public static final String GUIDE_CONFIG_FIELD = "field";
    public static final String GUIDE_CONFIG_FIELD_NAME = "name";
    public static final String GUIDE_CONFIG_FIELD_EXCEL = "excelTitle";
    public static final String GUIDE_CONFIG_FIELD_INDEX = "index";
    public static final String GUIDE_CONFIG_FIELD_DFVALUE = "defaultValue";
    public static final String GUIDE_CONFIG_FIELD_DATATYPE = "dataType";
    public static final String GUIDE_CONFIG_FIELD_FORMAT = "format";
    public static final String GUIDE_CONFIG_FIELD_NULLABLE = "nullable";
    public static final String GUIDE_CONFIG_FIELD_CONVERT = "convert";
    public static final String GUIDE_CONFIG_FIELD_IMPORTED = "imported";
    public static final String GUIDE_CONFIG_FIELD_ENTRY = "entry";
    public static final String GUIDE_CONFIG_ENTRY_KEY = "key";
    public static final String GUIDE_CONFIG_ENTRY_VALUE = "value";
    
    
    /**
     * 其他常量
     */
    public static final String DATA_TYPE_STRING = "String";
    public static final String DATA_TYPE_INTEGER = "Integer";
    public static final String DATA_TYPE_LONG = "Long";
    public static final String DATA_TYPE_DOUBLE = "Double";
    public static final String DATA_TYPE_DATE = "Date";
    public static final String DATA_TYPE_DATE_SYSDATE = "sysDate";
    
}
