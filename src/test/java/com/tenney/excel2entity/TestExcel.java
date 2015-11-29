/**
 * 版权所有：Nanjing Being information technology Co., LTD
 * 项目名称: tenney-excel2entity
 * 类名称:TestExcel.java
 * 包名称:com.tenney.excel2entity
 * 
 * 创建日期:2015年11月18日 
 * 创建人:唐雄飞		
 * <author>      <time>      <version>    <desc>
 * 唐雄飞     下午5:28:28     	V1.0        N/A
 */

package com.tenney.excel2entity;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import com.tenney.excel2entity.ExcelConstants.ImageType;
import com.tenney.excel2entity.lang.excel.ExcelBuilder;
import com.tenney.excel2entity.support.ICallBackMessage;
import com.tenney.excel2entity.support.IExcelReadCallBack;
import com.tenney.excel2entity.support.impl.TextAndStyleMessage;

/**
 *类说明:<br/>
 *创建人:唐雄飞<br/>
 *创建日期:2015年11月18日<br/> 
 * 
 */
public class TestExcel {
	
	@Test
	public void testEnum(){
		System.out.println(ImageType.JPG.name());
	}

	@Test
	public void testExport() throws Exception{
		String file = TestExcel.class.getResource("/excel-exp-imp.xml").getFile();
        ExcelGuideProvider provider = new ExcelGuideProvider(file);
        FileOutputStream fos = new FileOutputStream("/Users/tenney/Desktop/ExcelGuideProvider.xlsx");
        List<Map<String, Object>> dataSet = new ArrayList<>();
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("name", "小小船儿水中游");
        map.put("state", "");
        map.put("createDate", new Date());
        map.put("resType", "试题");
        map.put("thumb", new File("/Users/tenney/Pictures/11.jpg"));
        dataSet.add(map);
        
        map.put("name", "小小船儿水中游");
        map.put("state", "");
        map.put("createDate", new Date());
        map.put("resType", "试题");
        BufferedImage bufferImg = ImageIO.read(new File("/Users/tenney/Pictures/11.jpg"));
        map.put("thumb", bufferImg);
        dataSet.add(map);
        
        map.put("name", "小小船儿水中游");
        map.put("state", "");
        map.put("createDate", new Date());
        map.put("resType", "试题");
        BufferedImage bufferImg1 = ImageIO.read(new FileInputStream("/Users/tenney/Pictures/11.jpg"));
        map.put("thumb", bufferImg1);
        dataSet.add(map);
        
        provider.WriteToExcel(fos, "guideSimpleDemo", dataSet,"XLSX");
        
        System.out.println("导出成功->" + file);
	}
	
	@Test
	public void testImport() throws Exception{
		//导入数据测试
      ExcelGuideProvider excelGuideProvider = new ExcelGuideProvider("D:\\Intellij\\product\\branches\\eicsp\\grails-app\\conf\\resource\\excel-exp-imp.xml");
      FileOutputStream fos = new FileOutputStream("C:\\Users\\tenney\\Desktop\\导入结果.xls");
      File input = new File("C:\\Users\\tenney\\Desktop\\系统角色列表.xls");
      //正常导出
//      Collection<SysRole> roles = excelGuideProvider.readFromExcel(input, "sysRoleExportToEntity",fos);
//      System.out.println(roles);
      
      //回写导出测试
      Collection<Map<String,Object>> dataSet = excelGuideProvider.readFromExcel(new FileInputStream(input), "sysRoleExportToEntity",fos,new IExcelReadCallBack<Map<String,Object>>(){

		/**
		 * 方法描述:
		 * @see com.tenney.excel2entity.support.IExcelReadCallBack#mapRow(java.lang.Object, org.apache.poi.ss.usermodel.Workbook)
		 * @param voBean
		 * @param workbook
		 * @return
		 */
		@Override
		public ICallBackMessage mapRow(Map<String, Object> voBean, Workbook workbook) {
			System.out.println("回调了:" + voBean.get("name"));
            return new TextAndStyleMessage("成功",ExcelBuilder.buildMessageStyle(workbook));
		}
          
      });
	}

}
