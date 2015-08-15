/**
 * 版权所有：tenney
 * 项目名称: eicsp
 * 类名称:DateHelper.java
 * 包名称:com.tenney.excel2entity.lang
 * 
 * 创建日期:2013年10月21日 下午4:36:17
 * 创建人： 唐雄飞		
 * <修改人>      <时间>      <版本号>    <描述>
 * 唐雄飞      2013年10月21日     	V1.0.0        N/A
 */

package com.tenney.excel2entity.lang;

import org.apache.commons.lang3.StringUtils;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 类说明: <br/>
 * 创建人: 唐雄飞 <br/>
 * 创建日期:2013年10月21日 <br/> 
 * 
 */
public class DateHelper
{

    /**
     * 默认时间格式化格式
     */
    public static final String DEFAULT_DATE_FORMAT = "yyyy-MM-dd";
    /**
     * 日期格式化对象
     */
    public static final SimpleDateFormat SIMPE_DATA_FORMAT =  new SimpleDateFormat(DEFAULT_DATE_FORMAT);
    
    /**
     * 
     * 方法描述: 根据给定时间，格式返回解析的日期 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月21日 <br>
     * @param source
     * @param format
     * @return
     * @throws ParseException <br>
     */
    public static Date parse(String source,String format) throws ParseException{
        if(StringUtils.isBlank(source)){
            return null;
        }
        if(StringUtils.isNotBlank(format)){
            SIMPE_DATA_FORMAT.applyPattern(format);
            return SIMPE_DATA_FORMAT.parse(source);
        }else{
            return SIMPE_DATA_FORMAT.parse(source);
        }
    }
    
    /**
     * 
     * 方法描述: 根据给定日期，格式化时间 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月21日 <br>
     * @param source
     * @param format
     * @return <br>
     */
    public static String format(Date source,String format){
        if(null == source){
            return "";
        }
        if(StringUtils.isNotBlank(format)){
            SIMPE_DATA_FORMAT.applyPattern(format);
            return SIMPE_DATA_FORMAT.format(source);
        }else{
            return SIMPE_DATA_FORMAT.format(source);
        }
    }
}
