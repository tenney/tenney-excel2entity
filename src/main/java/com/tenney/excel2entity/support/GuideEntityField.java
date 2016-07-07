/**
 * 版权所有：tenney
 * 项目名称: eicsp
 * 类名称:GuideEntityField.java
 * 包名称:com.tenney.excel2entity.support
 * 
 * 创建日期:2013年10月18日 下午4:36:39
 * 创建人： 唐雄飞		
 * <修改人>      <时间>      <版本号>    <描述>
 * 唐雄飞      2013年10月18日     	V1.0.0        N/A
 */

package com.tenney.excel2entity.support;

import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;

import com.tenney.excel2entity.ExcelConstants;
import com.tenney.excel2entity.ExcelConstants.ImageType;

/**
 * 类说明: <br/>
 * 创建人: 唐雄飞 <br/>
 * 创建日期:2013年10月18日 <br/>
 * 
 */
public class GuideEntityField implements Comparable<GuideEntityField>
{

	private String name;//字段名
    private String excelTitle;//表格标题
    private Integer index = 0;//表格顺序
    private String dataType;//数据类型
    private String format;//日期转换格式
    private Boolean imported = true;//数据导入时，是否需要提供该字段
    private Boolean convert;//字段是否需要转换输出
    private Boolean nullable;//是否允许为空
    private String defaultValue;//默认值
    
    private Short widthOfColumn;  //X个字符的宽度
    private ExcelConstants.ImageType imageType = ImageType.PNG; //图片类型

    /**
     * 需要转换的子集
     */
    private Map<String, String> entrys = new HashMap<String, String>();

    /**
     * name的getter方法
     * 
     * @return the name
     */
    public String getName()
    {
        return name;
    }

    /**
     * name的setter方法
     * 
     * @param name the name to set
     */
    public void setName(String name)
    {
        this.name = name;
    }

    /**
     * excelTitle的getter方法
     * 
     * @return the excelTitle
     */
    public String getExcelTitle()
    {
        return excelTitle;
    }

    /**
     * excelTitle的setter方法
     * 
     * @param excelTitle the excelTitle to set
     */
    public void setExcelTitle(String excelTitle)
    {
        this.excelTitle = excelTitle;
    }

    /**
     * index的getter方法
     * 
     * @return the index
     */
    public Integer getIndex()
    {
        return index;
    }

    /**
     * index的setter方法
     * 
     * @param index the index to set
     */
    public void setIndex(Integer index)
    {
        this.index = index;
    }

    /**
     * dataType的getter方法
     * 
     * @return the dataType
     */
    public String getDataType()
    {
        return dataType;
    }

    /**
     * dataType的setter方法
     * 
     * @param dataType the dataType to set
     */
    public void setDataType(String dataType)
    {
        this.dataType = dataType;
    }

    /**
     * nullable的getter方法
     * 
     * @return the nullable
     */
    public Boolean getNullable()
    {
        return nullable;
    }

    /**
     * nullable的setter方法
     * 
     * @param nullable the nullable to set
     */
    public void setNullable(Boolean nullable)
    {
        this.nullable = nullable;
    }

    /**
     * defaultValue的getter方法
     * 
     * @return the defaultValue
     */
    public String getDefaultValue()
    {
        return defaultValue;
    }

    /**
     * defaultValue的setter方法
     * 
     * @param defaultValue the defaultValue to set
     */
    public void setDefaultValue(String defaultValue)
    {
        this.defaultValue = defaultValue;
    }

    /**
     * format的getter方法
     * 
     * @return the format
     */
    public String getFormat()
    {
        return format;
    }

    /**
     * format的setter方法
     * 
     * @param format the format to set
     */
    public void setFormat(String format)
    {
        this.format = format;
    }
    
    /**
	 * imported的getter方法
	 * @return the imported
	 */
	public Boolean getImported() {
		return imported;
	}

	/**
	 * imported的setter方法
	 * @param imported the imported to set
	 */
	public void setImported(Boolean imported) {
		this.imported = imported;
	}

	/**
     * convert的getter方法
     * 
     * @return the convert
     */
    public Boolean getConvert()
    {
        return convert;
    }

    /**
     * convert的setter方法
     * 
     * @param convert the convert to set
     */
    public void setConvert(Boolean convert)
    {
        this.convert = convert;
    }
    
	/**
	 * widthOfColumn的getter方法
	 * @return the widthOfColumn
	 */
	public Short getWidthOfColumn() {
		return widthOfColumn;
	}

	/**
	 * widthOfColumn的setter方法
	 * @param widthOfColumn the widthOfColumn to set
	 */
	public void setWidthOfColumn(Short widthOfColumn) {
		this.widthOfColumn = widthOfColumn;
	}

	/**
	 * imageType的getter方法
	 * @return the imageType
	 */
	public ExcelConstants.ImageType getImageType() {
		return imageType;
	}

	/**
	 * imageType的setter方法
	 * @param imageType the imageType to set
	 */
	public void setImageType(ExcelConstants.ImageType imageType) {
		this.imageType = imageType;
	}

	/**
     * entrys的getter方法
     * 
     * @return the entrys
     */
    public Map<String, String> getEntrys()
    {
        return entrys;
    }

    /**
     * entrys的setter方法
     * 
     * @param entrys the entrys to set
     */
    public void setEntrys(Map<String, String> entrys)
    {
        this.entrys = entrys;
    }

    /**
     * 方法描述:
     * @see java.lang.Comparable#compareTo(java.lang.Object)
     * @param o
     * @return
     */
    @Override
    public int compareTo(GuideEntityField entity)
    {
        if(this.getIndex() == null || entity.getIndex() == null || entity.getIndex().equals(this.getIndex())){
            return entity.getName().compareTo(this.getName());
        }
        return entity.getIndex() < this.getIndex() ? 1:-1;
    }

    /**
     * 方法描述:
     * @see java.lang.Object#toString()
     * @return
     */
    @Override
    public String toString()
    {
        return  "[name:" + this.getName() + " ,excelTitle: " + this.getExcelTitle() + " ,Index: " + this.getIndex()  + " , dataType :"+ this.getDataType() + ",nullable :" + this.getNullable()  + "]";
    }

    /**
     * 方法描述:
     * @see java.lang.Object#hashCode()
     * @return
     */
    @Override
    public int hashCode()
    {
        return super.hashCode();
    }

    /**
     * 方法描述:
     * @see java.lang.Object#equals(java.lang.Object)
     * @param obj
     * @return
     */
    @Override
    public boolean equals(Object obj)
    {
        return super.equals(obj);
    }
    
    /**
     * 
     * 方法描述: 根据值取键,用于Excel解析反转 <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月25日 <br>
     * @param eValue
     * @return <br>
     */
    public String getEntryKey(String eValue){
        if(this.entrys != null && StringUtils.isBlank(eValue)){
            for(String v : this.entrys.keySet()){
                if(eValue.equals(this.entrys.get(v))){
                    return v;
                }
            }
        }
        return null;
    }
}
