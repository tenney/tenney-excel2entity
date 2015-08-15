/**
 * 版权所有：tenney
 * 项目名称: eicsp
 * 类名称:ICallBackMessage.java
 * 包名称:com.tenney.excel2entity.support
 * 
 * 创建日期:2013年10月25日 下午8:57:05
 * 创建人： 唐雄飞		
 * <修改人>      <时间>      <版本号>    <描述>
 * 唐雄飞      2013年10月25日     	V1.0.0        N/A
 */

package com.tenney.excel2entity.support;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 类说明: <br/>
 * 创建人: 唐雄飞 <br/>
 * 创建日期:2013年10月25日 <br/> 
 * 
 */
public interface ICallBackMessage
{
    /**
     * 
     * 方法描述: 返回消息字符 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月25日 <br>
     * @return <br>
     */
    String getMessage();
    /**
     * 
     * 方法描述: 返回表格中的消息样式 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月25日 <br>
     * @return <br>
     */
    CellStyle getStyle();
    
}
