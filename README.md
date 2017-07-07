# ExcelKit

> 简单、好用且轻量级的海量Excel文件导入导出解决方案。



# 更新日志：

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
> ExcelKit-Example完整示例程序 ([https://github.com/wuwz/ExcelKit-Example][1])

![image](https://raw.githubusercontent.com/wuwz/ExcelKit-Example/master/example.gif)

# 如何使用？


1.引入Maven依赖或下载jar包

> 使用本工具无需关注poi依赖问题（只需引入以下相关jar包)。

``` xml
    <dependency>
        <groupId>org.wuwz</groupId>
        <artifactId>ExcelKit</artifactId>
        <version>${maven库上最新版本号https://mvnrepository.com/artifact/org.wuwz/ExcelKit}</version>
    </dependency>

		
    <!--以下视情况而定-->
    <dependency>
        <groupId>log4j</groupId>
        <artifactId>log4j</artifactId>
        <version>1.2.9</version>
    </dependency>
    <dependency>
        <groupId>javax.servlet</groupId>
        <artifactId>javax.servlet-api</artifactId>
        <version>3.0.1</version>
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
	    
	    @ExportConfig(value = "密码", replace = "******", color = HSSFColor.RED.index)
	    private String password;
	
	    @ExportConfig(value = "性别", width = 50, convert = "s:1=男,2=女")
	    private Integer sex;
	
	    @ExportConfig(value = "年级", convert = "c:org.wuwz.poi.test.GradeIdConvert")
	    private Integer gradeId;
	    
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

   需要实现org.wuwz.poi.convert.ExportConvert接口
   
   表达式: ```"c:org.wuwz.poi.test.GradeIdConvert"```
	

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

1.海量数据Excel文件读取（边读边处理）：

	

``` java
	ExcelKit.$Import()
	    .setEmptyCellValue(null) // 设置空单元格的值,默认为null,可不设置
	    .readExcel(new File("c:/bigexcel.xlsx"), new ReadHandler() {
	    
            @Override
            public void handler(int sheetIndex, int rowIndex, List<String> row) {
                if(rowIndex == 0) return; //排除第一行..
                
                    System.out.println("当前行："+rowIndex);
                    System.out.println("行数据："+row);
                    System.out.println();
                    
                    // 入库解析
                    if(row.get(0) != null) {
                        // UID...
                    } 
                    
                    if(row.get(1) != null) {
                        // 用户名..
                    }
            }
	    
	    });
```


 

2.生成Excel文件到本地：
 

``` java
	ExcelKit.$Builder(User.class)
	    .setMaxSheetRecords(10000) //设置每个sheet的最大记录数,默认为10000,可不设置
	    .toExcel(records, "用户数据", new FileOutputStream(new File("c:/test001.xlsx")));
```
	
	

  [1]: https://github.com/wuwz/ExcelKit-Example