/**
 * 版权所有：tenney
 * 项目名称: eicsp
 * 类名称:ExcelGuideException.java
 * 包名称:com.tenney.excel2entity.lang
 * 
 * 创建日期:2013年10月18日 下午4:54:07
 * 创建人： 唐雄飞		
 * <修改人>      <时间>      <版本号>    <描述>
 * 唐雄飞      2013年10月18日     	V1.0.0        N/A
 */

package com.tenney.excel2entity.lang;

/**
 * 类说明: <br/>
 * 创建人: 唐雄飞 <br/>
 * 创建日期:2013年10月18日 <br/>
 * 
 */
public class ExcelGuideException extends Exception
{

    /**
     * 变量类型：long 变量：serialVersionUID
     */
    private static final long serialVersionUID = 1L;

    public ExcelGuideException(Throwable cause)
    {
        super(cause);
    }

    /**
     * 变量类型：long 变量：serialVersionUID
     */

    /** 错误码 */
    protected String errorCode;

    /**
     * 构造函数
     * 
     * @param errorCode 错误码
     */
    public ExcelGuideException(String errorCode)
    {
        super(errorCode);
        this.errorCode = errorCode;
    }

    /**
     * 构造函数
     * 
     * @param errorCode 错误码
     * @param message 异常信息
     */
    public ExcelGuideException(String errorCode, String message)
    {
        super(message);
        this.errorCode = errorCode;
    }

    /**
     * 构造函数
     * 
     * @param errorCode 错误码
     * @param cause 原异常
     */
    public ExcelGuideException(String errorCode, Throwable cause)
    {
        super(errorCode, cause);
        this.errorCode = errorCode;
    }

    /**
     * 构造函数
     * 
     * @param errorCode 错误码
     * @param message 异常信息
     * @param cause 原异常
     */
    public ExcelGuideException(String errorCode, String message, Throwable cause)
    {
        super(message, cause);
        this.errorCode = errorCode;
    }

    /**
     * 获取错误码
     * 
     * @return 错误码
     */
    public String getErrorCode()
    {
        return errorCode;
    }

}
