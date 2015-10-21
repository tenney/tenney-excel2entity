# tenney-excel2entity
java对excel读写的封装,通过配置文件对表格进行读取和写入，具体请看配置文件

<?xml version="1.0" encoding="UTF-8"?>
<guides excel="xls">
	<!-- 
		注意： class的值如果为实体类，则该类型必须实现IExcelEntity接口
		导入导出均支持多种参数重载、采用范型和接口方式规范参数，
		其中导入数据支持返回Collection集合，也支持callback方式逐条数据处理，并将处理结果回写至表格，返回给用户
		导入导出均支持处理错误信息回写至表格。
		使用案例，请详细参考 TestMain
		
		不够完善的地方： 
			1、对于导入模板要求较严，尤其表头列的定义，不允许随意更改模板
			2、表格的读写依赖于POI插件
			3、依赖于spring的core包，用于集成在spring项目中时，支持classpath:*的方式加载配置文件
			4、暂时未做线程同步安全考虑、对于大量数据的处理应该有瓶颈，使用时，自行考虑数据量的问题。
	
		id		实体配置ID，需要保证其唯一性
		name	导出excel时的sheet名称，保存时的文件名
		class	转换记录数据载体，支持Map 和  domain
	 -->
	<entity id="guideSimpleDemo" name="导出配置测试" class="java.util.HashMap">
		<!-- 
			name			实体字段名/Map键名
			excelTile		Excel标题
			index			Excel字段顺序
			dataType		对应JAVA的数据类型，支持String,Long,Integer,Double,Date
			nullable		导入数据时，该字段是否允许为空，若为false,则最好指定 defaultValue的值
			defaultValue	导入数据时，若字段值为空时，默认取该值
			format			dataType为Date时，需要指定的格式化日期
			convert			是否需要字段转换 ,如 ：F/M >> 女/男 ,如果指定该值为true,则必须提供子元素集合
			imported		导入时是否需要提供数据  add by TXF 2013-11-04
		 -->
		<field name="name" excelTitle="资源标题" index="1" dataType="String" nullable="true" defaultValue=""></field>
		<field name="state" excelTitle="资源状态" index="2" dataType="Integer" nullable="false" defaultValue="1" convert="true">
			<entry key="0" value="无效"/>
			<entry key="1" value="有效"/>
		</field>
		<field name="createDate" excelTitle="上传时间" index="3" dataType="Date" format="yyyy年MM月dd日" nullable="false" defaultValue="sysdate"></field>
		<field name="resType" excelTitle="资源类型" index="4" dataType="String"  nullable="false" defaultValue="aa"></field>
		<field name="resSource" excelTitle="资源来源" index="4" dataType="String"  nullable="false" defaultValue="aa"></field>
	</entity>
</guides>
