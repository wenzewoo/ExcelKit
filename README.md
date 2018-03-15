# ExcelKit

> 简单、好用且轻量级的海量Excel文件导入导出解决方案。



# 更新日志：
*   2018-03-15 (version: 2.0.6)： 已经发布到Maven中央仓库(注意: GroupId以及包名修改为com.wuwenze), 等待中央仓库同步中..
``` xml
    <dependency>
        <groupId>com.wuwenze</groupId>
        <artifactId>ExcelKit</artifactId>
        <version>2.0.6</version>
    </dependency>
```

*   2018-03-11 (version: 2.0.6)：

    1. 修复导出文件后再读取该文件的数据错乱问题(数据翻倍). 
    
    2. 修复导出文件后再使用指定sheetIndex读取不到数据的问题 (感谢@OneToOne)
    
    3. convert转换器oldValue数据类型更改为Object, 用于支持其他类型的转换
    
    4. 提供完整的导出导入示例, 包含数据校验, 错误消息获取(详见: /src/test/java/...)
    
    5. 待解决: Maven中心仓库上传jar包不完整的问题(目前还没找到原因, 可能要换个账号发布了...)

*   2017-12-28：修复大文件excel内容错误问题, 取消字体颜色设置, 提升文件导出效率.

*   2017-9-14：下拉框，流上传，行尾空单元格，导出单元格格式（感谢@maxcess）

*   2017-9-11：修复：补全行尾可能缺失的单元格，待修复: 补全行首可能缺失的单元格

*   2017-7-8： 修复读取Excel文件后，临时文件一直占用的问题

*   2017-6-25： 修复解决导入以及导出单元格空值的问题

*   2017-4-11： 代码重构,不再兼容1.x。
*   1. @ExportConfig重构,支持convert简单转换器.
*   2. Excel读取性能优化,由usermodel包下的API重构为eventusermode包下的API,以SAX的方式,边读边处理,性能得到了极大的提升,可轻松实现百万级别的海量数据处理.
*   3. 写Excel文件的API重构为streaming包下的API，同时支持多sheet导出.
*   4. 取消了对Excel2003以及以下版本的excel文件支持.
*   5. 包结构重构,便于以后的版本扩展和升级,不再兼容1.x版本.


# 编译环境：
> 使用``` jdk1.6.0_45 ``` 和```maven-3.2.5```进行项目构建,理论上支持```jdk6+```。

# 使用效果：
> 示例代码见: \ExcelKit\src\test\java\org\wuwz\poi\test\Demo.java

#### WEB环境下的导入导出
![image](https://raw.githubusercontent.com/wuwz/ExcelKit/master/example1.gif)

#### 极端环境下的性能测试(可用内存12m, 执行10000行数据的导出和导入操作)
![image](https://raw.githubusercontent.com/wuwz/ExcelKit/master/example2.gif)

# 如何使用？


1.引入Maven依赖或下载jar包

> 使用本工具无需关注poi依赖问题（只需引入以下相关jar包)。


``` xml
    <dependency>
        <groupId>com.wuwenze</groupId>
        <artifactId>ExcelKit</artifactId>
        <version>2.0.6</version>
    </dependency>
```

注：非maven项目如何下载jar包？参考：http://blog.csdn.net/itmyhome1990/article/details/50233773

> ExcelKit集成的jar:
``` xml
	<poi-version>3.8</poi-version>
	<beanutils-version>1.9.3</beanutils-version>
	<dom4j-version>1.6.1</dom4j-version>
	<jaxen-version>1.1.6</jaxen-version>
	<xerces-version>2.11.0</xerces-version>
	<xml-apis-version>1.4.01</xml-apis-version>
```

       

2.导出项配置（通过注解）：
 
``` java

	public class User {
	
	    @ExportConfig("UID")
	    private Integer uid;
	    
	    @ExportConfig("用户名")
	    private String username;
	    
	    @ExportConfig(value = "密码", replace = "******")
	    private String password;
	
	    @ExportConfig(value = "性别", width = 50, convert = "s:1=男,2=女")
	    private Integer sex;
	
	    @ExportConfig(value = "年级", convert = "c:com.wuwenze.poi.test.convert.GradeIdConvert")
	    private Integer gradeId;
	    
	    @ExportConfig(value = "下拉框", range="c:com.wuwenze.poi.test.convert.RangeConvert")
		private String gendex;
	    
	    // getter setter...
	    
	 }
```

2.1.实现ExportConvert（如果配置了convert属性,参考gradeId字段）：

> convert属性说明:

将单元格值进行转换后再导出：
目前支持以下几种场景：

1. 固定的数值转换为字符串值（如：1代表男，2代表女）

	表达式: ```"s:1=男,2=女"```
	
	
2. 数值对应的值需要查询数据库才能进行映射 (用的机会不多,一般这种情况直接在DAO处理)

   需要实现com.wuwenze.poi.convert.ExportConvert接口
   
   表达式: ```"c:com.wuwenze.poi.test.convert.GradeIdConvert"```
	

``` java

	public class GradeIdConvert implements ExportConvert {
	
	    static Map<Integer, String> records;
	    static {
	        //默认数据库字典查询 select * from tb_grades
	        records = new HashMap<Integer, String>();
	        records.put(1, "一年级学生");
	        records.put(2, "二年级学生");
	        records.put(3, "三年级学生");
	    }
	    
	    @Override
	    public String handler(Integer val) {
	        return records.get(val) != null ? records.get(val) : "无记录";
	    }
	
	}
```

        

3.一行代码执行浏览器导出：
> 注意：导出数量有内存限制，这取决于你的设备可用内存、JVM内存，建议适当的调大一点

``` java
	@RequestMapping("/export");
	public void export(HttpServletResponse response) {
	    List<User> users = dao.getUsers();
	    // 生成Excel并使用浏览器下载
	    ExcelKit.$Export(User.class, response).toExcel(users, "用户信息");
	}
```

		

	

# 常用例子：

1.海量数据Excel文件读取、导入（边读边处理）：

	

``` java
		/*1. 海量数据导入 （伪代码）*/
		
		// 获取上传文件
		File upload = null;
		
		// 读取并解析文件
		final List<Object> exportData = new ArrayList<Object>();
		ExcelKit.$Import().readExcel(upload, new ReadHandler() {
			
			@Override
			public void handler(int sheetIndex, int rowIndex, List<String> row) {
				try {
					// 排除表头
					if(rowIndex == 0) return;
					
					// 验证行数据
					if(!validRow(row)) {
						// TODO: 行数据rowIndex验证失败，记录
						
					} else {
						// 解析行数据
						
						// 方案1：记录行数据，读取完成后批量入库
						exportData.add(analysisRow(row));
						
						// 方案2：单行直接入库
						// xxx.save(analysisRow(row));
					} 
					
				} catch (Exception e) {
					e.printStackTrace();
					// 读取行：rowIndex发生异常，记录
				}
			}

			private Object analysisRow(List<String> row) {
				// TODO 解析行数据为对象
				return null;
			}

			private boolean validRow(List<String> row) {
				// TODO 验证行数据
				return true;
			}
		});
		
		// 方案1： 文件读取解析完毕, 批量入库。
		xxx.batchSave(exportData);
		
		// TODO 响应结果，成功条数、失败条数（行、原因）
```


 

2.生成Excel文件到本地：
 

``` java
	ExcelKit.$Builder(User.class)
	    .setMaxSheetRecords(10000) //设置每个sheet的最大记录数,默认为10000,可不设置
	    .toExcel(records, "用户数据", new FileOutputStream(new File("c:/test001.xlsx")));
```
	
	

  [1]: https://github.com/wuwz/ExcelKit-Example