/**
 * 版权所有：tenney
 * 项目名称: excel-guide-plugin
 * 类名称:XmlParser.java
 * 包名称:com.tenney.excel2entity.lang.xml
 * 
 * 创建日期:2013年10月18日 下午2:28:42
 * 创建人： 唐雄飞		
 * <修改人>      <时间>      <版本号>    <描述>
 * 唐雄飞      2013年10月18日     	V1.0.0        N/A
 */

package com.tenney.excel2entity.lang.xml;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.io.SAXReader;

import java.io.InputStream;

/**
 * 类说明: <br/>
 * 创建人: 唐雄飞 <br/>
 * 创建日期:2013年10月18日 <br/> 
 * 
 */
public class XmlParser
{
    /**
     * 
     * 方法描述: 读取XML文档 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月18日 <br>
     * @param source
     * @return
     * @throws DocumentException <br>
     */
    public static Document parseFromSource(InputStream source) throws DocumentException{
        return new SAXReader().read(source);
    }
    
    /**
     * 
     * 方法描述: 读取XML[重载] <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月18日 <br>
     * @param filePath
     * @return
     * @throws DocumentException <br>
     */
    public static Document parseFromSource(String filePath) throws DocumentException{
        return new SAXReader().read(filePath);
    }
}
