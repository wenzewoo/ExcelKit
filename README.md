# ExcelKit

> 简单,好用且轻量级Excel文件导入导出工具。


# 如何使用？

 1. 引入Maven依赖：
 

        <dependency>
			<groupId>org.wuwz</groupId>
			<artifactId>ExcelKit</artifactId>
			<version>1.0</version>
		</dependency>

        <!--以下视情况而定-->
		<dependency>
			<groupId>javax</groupId>
			<artifactId>javaee-api</artifactId>
			<version>7.0</version>
		</dependency>
		<dependency>
			<groupId>javax.servlet</groupId>
			<artifactId>javax.servlet-api</artifactId>
			<version>3.1.0</version>
		</dependency>

 2. 导出项配置（通过注解）：
 

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

 3. 一行代码导出：
 

		List<User> users = dao.getUsers();
		
		// 生成Excel并使用浏览器下载
		ExcelKit.$Export(User.class, response).toExcel(users, "用户信息");


# 其他使用例子

 1. 导入Excel读取数据：

    	List<User> users = Lists.newArrayList();
		
		//导入数据。
		File excelFile = new File("C:\\Users\\Administrator\\Desktop\\excel.xlsx");
		List<List<String>> excelDatas = ExcelKit.$Import().readExcel(excelFile);
		for (List<String> ed : excelDatas) {
			
			User u = new User();
			u.setUid(Integer.valueOf(ed.get(0)));
			u.setUsername(ed.get(1));
			u.setPassword(ed.get(2));
			u.setNickname(ed.get(3));
			
			u.setAge(18);
			users.add(u);
		}
		
		System.out.println(users);

 

 2. 生成Excel文件到本地、生成导入模版文件：
 

        // 生成本地文件
		File excelFile = new File("C:\\Users\\Administrator\\Desktop\\excel.xlsx");
		ExcelKit.$Builder(User.class).toExcel(users, "用户信息", new FileOutputStream(excelFile));
		
		// 生成Excel导入模版文件。
		users.clear();
		File templateFile = new File("C:\\Users\\Administrator\\Desktop\\import_template.xlsx");
		ExcelKit.$Builder(User.class).toExcel(users, "用户信息", new FileOutputStream(templateFile));