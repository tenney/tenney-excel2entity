package com.tenney.excel2entity.support;

import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;

public class GuideTitle implements Comparable<GuideTitle>{

	private String name;//字段名
	private String excelTitle;//表格标题
	private Boolean imported = false;//数据导入时，是否需要提供该字段
	private Integer rowspan;//合并单元格行
	private Integer colspan;//合并单元格列
	private Map<String, String> entrys = new HashMap<String, String>();
	/**
	 *以下字段暂时保留 
	 * */
//	private String mergeC;//合并单元格行
//	private String mergeR;//合并单元格列
//	private String fontColor;//字体颜色
//	private String fontSize;//字体大小
//	private String bgColor;//背景色
	

	
	
	 /**
     * 方法描述:
     * @see java.lang.Object#toString()
     * @return
     */
    @Override
    public String toString()
    {
        return  "[name:" + this.getName() + " ,excelTitle: " + this.getExcelTitle() + " , rowspan :"+ this.getRowspan() + ",colspan :" + this.getColspan()  + "]";
    }

    /**
	 * @return the name
	 */
	public String getName() {
		return name;
	}

	/**
	 * @param name the name to set
	 */
	public void setName(String name) {
		this.name = name;
	}

	/**
	 * @return the excelTitle
	 */
	public String getExcelTitle() {
		return excelTitle;
	}

	/**
	 * @param excelTitle the excelTitle to set
	 */
	public void setExcelTitle(String excelTitle) {
		this.excelTitle = excelTitle;
	}

	/**
	 * @return the imported
	 */
	public Boolean getImported() {
		return imported;
	}

	/**
	 * @param imported the imported to set
	 */
	public void setImported(Boolean imported) {
		this.imported = imported;
	}


	public Integer getRowspan() {
		return rowspan;
	}

	public void setRowspan(Integer rowspan) {
		this.rowspan = rowspan;
	}

	public Integer getColspan() {
		return colspan;
	}

	public void setColspan(Integer colspan) {
		this.colspan = colspan;
	}

	/**
	 * @return the entrys
	 */
	public Map<String, String> getEntrys() {
		return entrys;
	}

	/**
	 * @param entrys the entrys to set
	 */
	public void setEntrys(Map<String, String> entrys) {
		this.entrys = entrys;
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

	@Override
	public int compareTo(GuideTitle o) {
		// TODO Auto-generated method stub
		return 0;
	}
}
