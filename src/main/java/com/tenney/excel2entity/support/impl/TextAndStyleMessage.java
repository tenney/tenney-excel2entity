/**
 * 版权所有：tenney
 * 项目名称: eicsp
 * 类名称:TextAndStyleMessage.java
 * 包名称:com.tenney.excel2entity.support.impl
 * 
 * 创建日期:2013年10月25日 下午9:22:47
 * 创建人： 唐雄飞		
 * <修改人>      <时间>      <版本号>    <描述>
 * 唐雄飞      2013年10月25日     	V1.0.0        N/A
 */

package com.tenney.excel2entity.support.impl;

import com.tenney.excel2entity.support.ICallBackMessage;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 类说明: <br/>
 * 创建人: 唐雄飞 <br/>
 * 创建日期:2013年10月25日 <br/> 
 * 
 */
public class TextAndStyleMessage implements ICallBackMessage
{

    /**
     * 回写的消息
     */
    private String message;
    private CellStyle style;
    
    public TextAndStyleMessage(String message,CellStyle style){
        this.message = message;
        this.style = style;
    }
    /**
     * 方法描述:
     * @see com.tenney.excel2entity.support.ICallBackMessage#getMessage()
     * @return
     */
    @Override
    public String getMessage()
    {
        return this.message;
    }

    /**
     * 方法描述:
     * @see com.tenney.excel2entity.support.ICallBackMessage#getStyle()
     * @return
     */
    @Override
    public CellStyle getStyle()
    {
        return this.style;
    }

}
