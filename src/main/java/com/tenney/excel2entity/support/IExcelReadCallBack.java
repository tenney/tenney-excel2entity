/**
 * 版权所有：tenney
 * 项目名称: eicsp
 * 类名称:IExcelReadCallBack.java
 * 包名称:com.tenney.excel2entity.support
 * 
 * 创建日期:2013年10月25日 下午8:37:32
 * 创建人： 唐雄飞		
 * <修改人>      <时间>      <版本号>    <描述>
 * 唐雄飞      2013年10月25日     	V1.0.0        N/A
 */

package com.tenney.excel2entity.support;

import org.apache.poi.ss.usermodel.Workbook;

import com.tenney.excel2entity.support.ICallBackMessage;

/**
 * 类说明: 行数据读取结果回写接口,用于将读取Excel数据的操作结果回写到Excel中 <br/>
 * 创建人: 唐雄飞 <br/>
 * 创建日期:2013年10月25日 <br/>
 * 
 */
public interface IExcelReadCallBack<T>
{
    /**
     * 
     * 方法描述: 每读取一行数据，将回调该方法 <br>
     * 创建人: 唐雄飞 <br>
     * 创建日期:2013年10月25日 <br>
     * 
     * @param voBean 该行数据封装后的数据载体
     * @param workbook 当前操作表格对象
     * @return 回写给excel表格的信息 <br>
     */
    ICallBackMessage mapRow(T voBean, Workbook workbook);
}
