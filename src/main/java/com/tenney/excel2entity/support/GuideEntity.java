/**
 * 版权所有：tenney
 * 类名称:GuideEntity.java
 * 包名称:com.tenney.excel2entity.support
 * 
 * 创建日期:2013年10月18日 下午3:21:53
 * 创建人： 唐雄飞		
 * <修改人>      <时间>      <版本号>    <描述>
 * 唐雄飞      2013年10月18日     	V1.0.0        N/A
 */

package com.tenney.excel2entity.support;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

/**
 * 类说明: <br/>
 * 创建人: 唐雄飞 <br/>
 * 创建日期:2013年10月18日 <br/>
 * 
 */
public class GuideEntity
{
    private String eid;//实体ID
    private String eName;//实体名
    private String entityClass;//实体对象类
    
    private Float heightOfRows = -1f; //行高，像素点， -1代表默认

    private GuideTitle title;
    /**
     * 实体对象字段集_TreeSet提供Comparable按字段的index进行排序 
     */
    private Set<GuideEntityField> fields = new TreeSet<GuideEntityField>();

    /**
     * GuideEntity.java的构造函数
     */
    public GuideEntity()
    {
        super();
    }

    /**
     * GuideEntity.java的构造函数
     * 
     * @param eid
     * @param eName
     * @param clazz
     * @param fields
     */
    public GuideEntity(String eid, String eName, String entityClass, Set<GuideEntityField> fields,GuideTitle title)
    {
        super();
        this.eid = eid;
        this.eName = eName;
        this.entityClass = entityClass;
        this.fields = fields;
        this.title = title;
    }

    /**
     * eid的getter方法
     * 
     * @return the eid
     */
    public String getEid()
    {
        return eid;
    }

    /**
     * eid的setter方法
     * 
     * @param eid the eid to set
     */
    public void setEid(String eid)
    {
        this.eid = eid;
    }

    /**
     * eName的getter方法
     * 
     * @return the eName
     */
    public String geteName()
    {
        return eName;
    }

    /**
     * eName的setter方法
     * 
     * @param eName the eName to set
     */
    public void seteName(String eName)
    {
        this.eName = eName;
    }

    /**
     * entityClass的getter方法
     * @return the entityClass
     */
    public String getEntityClass()
    {
        return entityClass;
    }

    /**
     * entityClass的setter方法
     * @param entityClass the entityClass to set
     */
    public void setEntityClass(String entityClass)
    {
        this.entityClass = entityClass;
    }
    
    /**
	 * heightOfRows的getter方法
	 * @return the heightOfRows
	 */
	public Float getHeightOfRows() {
		return heightOfRows;
	}

	/**
	 * heightOfRows的setter方法
	 * @param heightOfRows the heightOfRows to set
	 */
	public void setHeightOfRows(Float heightOfRows) {
		this.heightOfRows = heightOfRows;
	}

	/**
     * fields的getter方法
     * 
     * @return the fields
     */
    public Set<GuideEntityField> getFields()
    {
        return fields;
    }

    /**
     * fields的setter方法
     * 
     * @param fields the fields to set
     */
    public void setFields(Set<GuideEntityField> fields)
    {
        this.fields = fields;
    }

    /**
	 * @return the title
	 */
	public GuideTitle getTitle() {
		return title;
	}

	/**
	 * @param title the title to set
	 */
	public void setTitle(GuideTitle title) {
		this.title = title;
	}

	/**
     * 
     * 方法描述: 获取所有字段集合，用于Excel根据列名取字段<br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月25日 <br>
     * @return <br>
     */
    public Map<String,GuideEntityField> getFieldMap(){
        Map<String,GuideEntityField> maps = new HashMap<String,GuideEntityField>();
        if(this.fields != null){
            for(GuideEntityField field:this.fields){
                maps.put(field.getExcelTitle().trim(), field);
            }
        }
        return maps;
    }
    
    public String getGuideTitle(){
    	if(this.fields != null){
    		return title.getExcelTitle().trim();
    	}
    	return null;
    }
    
}
