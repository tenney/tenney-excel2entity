/**
 * 版权所有：tenney
 * 项目名称: eicsp
 * 类名称:TestMain.java
 * 包名称:com.tenney.excel2entity.test
 * 
 * 创建日期:2013年10月18日 下午5:08:48
 * 创建人： 唐雄飞		
 * <修改人>      <时间>      <版本号>    <描述>
 * 唐雄飞      2013年10月18日     	V1.0.0        N/A
 */

package com.tenney.excel2entity;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;

import org.dom4j.DocumentException;

import com.tenney.excel2entity.ExcelGuideProvider;
import com.tenney.excel2entity.lang.ExcelGuideException;

/**
 * 类说明: <br/>
 * 创建人: 唐雄飞 <br/>
 * 创建日期:2013年10月18日 <br/> 
 * 
 */
public class TestMain
{
    /**
     * 方法描述: <br>
     * 创建人: 唐雄飞  <br>
     * 创建日期:2013年10月18日 <br>
     * @param args <br>
     * @throws IOException 
     * @throws DocumentException 
     * @throws ExcelGuideException 
     * @throws NoSuchMethodException 
     * @throws InvocationTargetException 
     * @throws IllegalAccessException 
     */
    @SuppressWarnings("unchecked")
	public static void main(String[] args) throws DocumentException, IOException, ExcelGuideException, IllegalAccessException, InvocationTargetException, NoSuchMethodException
    {
//        导出数据测试
//        ExcelGuideProvider provider = new ExcelGuideProvider(TestMain.class.getResource("/excel-exp-imp.xml").getFile());
//        FileOutputStream fos = new FileOutputStream("d:\\ExcelGuideProvider.xlsx");
//        List dataSet = new ArrayList();
//        Map map = new HashMap();
//        map.put("name", "小小船儿水中游");
//        map.put("state", "");
//        map.put("createDate", new Date());
//        map.put("resType", "试题");
//        dataSet.add(map);
//        provider.WriteToExcel(fos, "guideSimpleDemo", dataSet,"XLS");
//        
        //导入数据测试
        ExcelGuideProvider excelGuideProvider = new ExcelGuideProvider("D:\\Intellij\\product\\branches\\eicsp\\grails-app\\conf\\resource\\excel-exp-imp.xml");
        FileOutputStream fos = new FileOutputStream("C:\\Users\\tenney\\Desktop\\导入结果.xls");
        File input = new File("C:\\Users\\tenney\\Desktop\\系统角色列表.xls");
        //正常导出
//        Collection<SysRole> roles = excelGuideProvider.readFromExcel(input, "sysRoleExportToEntity",fos);
//        System.out.println(roles);
        
        //回写导出测试
//        excelGuideProvider.readFromExcel(new FileInputStream(input), "sysRoleExportToEntity",fos,new IExcelReadCallBack<SysRole>(){
//            @Override
//            public ICallBackMessage mapRow(SysRole voBean, HSSFWorkbook workbook)
//            {
////                voBean.save(); //数据入库操作
////                return new SimpleCallBackTextMessage("导入成功!");
//                System.out.println("回调了:" + voBean.getName());
//                return new TextAndStyleMessage("成功",ExcelBuilder.buildMessageStyle(workbook));
//            }
//        });
    }
}
