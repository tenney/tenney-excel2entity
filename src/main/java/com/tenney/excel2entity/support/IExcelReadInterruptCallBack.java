/**
 * 版权所有：Nanjing Being information technology Co., LTD
 * 项目名称: tenney-excel2entity
 * 类名称:IExcelReadInterruptCallBack.java
 * 包名称:com.tenney.excel2entity.support
 * 
 * 创建日期:2015年12月8日 
 * 创建人:唐雄飞		
 * <author>      <time>      <version>    <desc>
 * 唐雄飞     下午9:47:21     	V1.0        N/A
 */

package com.tenney.excel2entity.support;

import org.apache.poi.ss.usermodel.Workbook;

/**
 *类说明:<br/>
 *创建人:唐雄飞<br/>
 *创建日期:2015年12月8日<br/> 
 * @param <T>
 * 
 */
public interface IExcelReadInterruptCallBack<T> extends IExcelReadCallBack<T> {
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
    ICallBackMessage mapRow(T voBean, Workbook workbook) throws Exception;
}
