# ExcelKit

> 简单,好用且轻量级Excel文件导入导出工具。

# 编译环境：
> 使用``` jdk1.6.0_45 ``` 和```maven-3.2.5```进行项目构建,理论上支持```jdk6+```。

# 使用效果：
> ExcelKit-Example完整示例程序 ([https://github.com/wuwz/ExcelKit-Example][1])
![image](https://raw.githubusercontent.com/wuwz/ExcelKit-Example/master/example.gif)

# 如何使用？


1.引入Maven依赖或下载jar包([点我下载ExcelKit-1.0.jar][2])

> 使用本工具无需关注poi依赖问题（只需引入以下相关jar包),完整的依赖说明见  ``` org.wuwz.poi.ExcelKit ``` 类注释。

``` xml
         <dependency>
			<groupId>org.wuwz</groupId>
			<artifactId>ExcelKit</artifactId>
			<version>1.0</version>
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

       

2.导出项配置（通过注解）：
 
``` java
	public class User {

        	@ExportConfig(value = "UID", width = 150)
        	private Integer uid;
        
        	@ExportConfig(value = "用户名", width = 200)
        	private String username;
        
        	@ExportConfig(value = "密码(不可见)", width = 120, isExportData = false)
        	private String password;
        
        	@ExportConfig(value = "昵称", width = 200)
        	private String nickname;
        
        	private Integer age;
        
        	// getter setter...
        }
```


        

3.一行代码执行浏览器导出：

``` java
	@RequestMapping("/export");
	public void export(HttpServletResponse response) {
		List<User> users = dao.getUsers();
		
		// 生成Excel并使用浏览器下载
		ExcelKit.$Export(User.class, response).toExcel(users, "用户信息");
	}
```

		

	

# 常用例子：

1.导入Excel读取数据：

	

``` java
	final List<User> users = Lists.newArrayList();
	
	//导入数据。
	File excelFile = new File("C:\\Users\\Administrator\\Desktop\\excel.xlsx");
	ExcelKit.$Import().readExcel(excelFile, new OnReadDataHandler() {
		
		@Override
		public void handler(List<String> rowData) {
			User u = new User();
			u.setUid(Integer.valueOf(rowData.get(0)));
			u.setUsername(rowData.get(1));
			u.setPassword(rowData.get(2));
			u.setNickname(rowData.get(3));
			
			u.setAge(18);
			users.add(u);
			
		}
	});
	
	System.out.println(users);
```


 

2.生成Excel文件到本地、生成导入模版文件：
 

	

``` java
	// 生成本地文件
	File excelFile = new File("C:\\Users\\Administrator\\Desktop\\excel.xlsx");
	ExcelKit.$Builder(User.class).toExcel(users, "用户信息", new FileOutputStream(excelFile));
	
	// 生成Excel导入模版文件。
	users.clear();
	File templateFile = new File("C:\\Users\\Administrator\\Desktop\\import_template.xlsx");
	ExcelKit.$Builder(User.class).toExcel(users, "用户信息", new FileOutputStream(templateFile));
```

		
        
# 其他例子（所有）：

> 本示例只做参考,可能不能直接运行（需要数据支持）。

``` java

    File excelFile = new File("C:\\Users\\Administrator\\Desktop\\excel.xlsx");
	
	//1. 生成本地文件
	ExcelKit.$Builder(User.class).toExcel(Db.getUsers(), "用户信息", new FileOutputStream(excelFile));
	
	//2. 文件导入模版
	ExcelKit.$Builder(User.class).toExcel(null, "用户信息", new FileOutputStream(excelFile));
	
	
	//3. 自定义Excel文件生成/导出(ExcelKit.$Export(class,response))
	ExcelKit.$Builder(User.class).toExcel(Db.getUsers(), "用户信息", ExcelType.EXCEL2007, new OnSettingHanlder() {
		
		@Override
		public CellStyle getHeadCellStyle(Workbook wb) {
			// 设置表头样式
			CellStyle cellStyle = wb.createCellStyle();
			Font font = wb.createFont();
			cellStyle.setAlignment(CellStyle.ALIGN_LEFT);// 对齐
			cellStyle.setFillForegroundColor(HSSFColor.GREEN.index);
			cellStyle.setFillBackgroundColor(HSSFColor.GREEN.index);
			font.setBoldweight(Font.BOLDWEIGHT_NORMAL);
			font.setFontHeightInPoints((short) 14);// 字体大小
			font.setColor(HSSFColor.WHITE.index);
			cellStyle.setFont(font);
			//......
			return cellStyle;
		}
		
		@Override
		public String getExportFileName(String sheetName) {
			// 设置导出文件名
			return String.format("导出-%s-%s", sheetName,System.currentTimeMillis());
		}
		
		@Override
		public CellStyle getBodyCellStyle(Workbook wb) {
			return null;
		}
	}, new FileOutputStream(excelFile));
	
	
	//4. 读取指定sheetIndex
	ExcelKit.$Import().readExcel(excelFile, 0, new OnReadDataHandler() {
		
		@Override
		public void handler(List<String> rowData) {
			
		}
	});
	
	//5. 设置空单元格值,默认为：“EMPTY_CELL_VALUE”
	final String emptyValue = "null";
	ExcelKit.$Import().setEmptyCellValue(emptyValue).readExcel(excelFile, new OnReadDataHandler() {
		
		@Override
		public void handler(List<String> rowData) {
			if(emptyValue.equals(rowData.get(0))) {
				//此单元格的值为空,需要额外处理
			}
		}
	});
	
	
	//6. 读取指定sheet、行、单元格
	int sheetIndex = 0;
	int startRowIndex = 0;
	int endRowIndex = 9;
	int startCellIndex = 0;
	int endCellIndex = 3;
	ExcelKit.$Import().readExcel(excelFile, new OnReadDataHandler() {
		
		@Override
		public void handler(List<String> rowData) {
			// TODO Auto-generated method stub
			
		}
	}, sheetIndex, startRowIndex, endRowIndex, startCellIndex, endCellIndex);
	
	
	//.....
```
		
		
		


  [1]: https://github.com/wuwz/ExcelKit-Example
  [2]: https://github.com/wuwz/ExcelKit/blob/master/ExcelKit-1.0.jar?raw=true