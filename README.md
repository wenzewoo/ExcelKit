# ExcelKit

> 简单、好用且轻量级的海量Excel文件导入导出解决方案。<反馈交流 QQ 群: 728244066>

***注意: 当前版本为全新重构的新版本, 不兼容老版本, 新版正在测试阶段, 请勿直接用于生产环境!***

## POM.xml 
```xml
<dependency>
    <groupId>com.wuwenze</groupId>
    <artifactId>ExcelKit</artifactId>
    <version>2.0.7</version>
</dependency>
```

## 示例
### 1. ExcelMapping

> 现支持两种配置方式: 注解 或者 XML
```java
@Data
@Builder
@ToString
@AllArgsConstructor
@NoArgsConstructor
@Excel("用户信息")
public class User {

  @ExcelField("编号")
  private Integer userId;

  @ExcelField(//
      value = "用户名",//
      required = true,//
      validator = UsernameValidator.class,//
      comment = "请填写用户名，最大长度为12，且不能重复"
  )
  private String username;

  @ExcelField(value = "密码", required = true, maxLength = 32)
  private String password;

  @ExcelField(value = "邮箱", validator = UserEmailValidator.class)
  private String email;

  @ExcelField(//
      value = "性别",//
      readConverterExp = "未知=0,男=1,女=2",//
      writeConverterExp = "0=未知,1=男,2=女",//
      options = SexOptions.class//
  )
  private Integer sex;

  @ExcelField(//
      value = "用户组",//
      name = "userGroup.name",//
      options = UserGroupNameOptions.class
  )
  private UserGroup userGroup;

  @ExcelField(value = "创建时间", dateFormat = "yyyy/MM/dd HH:mm:ss")
  private Date createAt;

  @ExcelField(//
      value = "自定义字段",//
      maxLength = 80,//
      comment = "可以乱填，但是长度不能操作80，导入时最终会转换为数字",//
      writeConverter = CustomizeFieldWriteConverter.class,// 写文件时，将数字转字符串
      readConverter = CustomizeFieldReadConverter.class// 读文件时，将字符串转数字
  )
  private Integer customizeField;
}
```

> XML 配置方式, 需将 xml 文件放置在classpath:excel-mapping/{entityClassName}.xml

```java
@Data
@Builder
@ToString
@AllArgsConstructor
@NoArgsConstructor
public class User2 {

  private Integer userId;
  private String username;
  private String password;
  private String email;
  private Integer sex;
  private UserGroup userGroup;
  private Date createAt;
  private Integer customizeField;
}
```

> classpath:excel-mapping/com.wuwenze.entity.User2.xml
```xml
<?xml version="1.0" encoding="UTF-8"?>
<excel-mapping name="用户信息">
  <property name="userId" column="编号"/>
  <property
    name="username"
    column="用户名"
    required="true"
    validator="com.wuwenze.validator.UsernameValidator"
    comment="请填写用户名，最大长度为12，且不能重复"
  />
  <property name="password" column="密码" required="true" maxLength="32"/>
  <property name="email" column="邮箱" validator="com.wuwenze.validator.UserEmailValidator"/>
  <property
    name="sex"
    column="性别"
    readConverterExp="未知=0,男=1,女=2"
    writeConverterExp="0=未知,1=男,2=女"
    options="com.wuwenze.options.SexOptions"
  />
  <property name="userGroup.name" column="用户组" options="com.wuwenze.options.UserGroupNameOptions" />
  <property name="createAt" column="创建时间" dateFormat="yyyy/MM/dd HH:mm:ss" />
  <property
    name="customizeField"
    column="自定义字段"
    maxLength="80"
    comment="可以乱填，但是长度不能操作80，导入时最终会转换为数字"
    writeConverter="com.wuwenze.converter.CustomizeFieldWriteConverter"
    readConverter="com.wuwenze.converter.CustomizeFieldReadConverter"
  />
</excel-mapping>
```
以上两种方式2选一即可, ExcelKit 会优先装载 XML 配置文件.


### 2. 一行代码构建 Excel 导入模板
> 使用 ExcelKit 提供的API 构建导入模板, 会根据配置生成批注, 下拉框等
``` java
// 生成导入模板（含3条示例数据）
@RequestMapping(value = "/downTemplate", method = RequestMethod.GET)
public void downTemplate(HttpServletResponse response) {
  List<User> userList = DbUtil.getUserList(3);
  ExcelKit.$Export(User.class, response).downXlsx(userList, true);
}
```

### 3. 执行文件导入
> 使用边读边处理的方式, 无需担心内存溢出, 也不用理会 Excel 文件到底有多大.

``` java
@RequestMapping(value = "/importUser", method = RequestMethod.POST)
public ResponseEntity<?> importUser(@RequestParam MultipartFile file)
    throws IOException {
  List<User> userList = Lists.newArrayList();
  Map<Integer, List<ExcelErrorField>> errorMap = Maps.newHashMap();
  ExcelKit.$Import(User.class)
      .readXlsx(file.getInputStream(), new ExcelReadHandler<User>() {

        @Override
        public void onSuccess(int sheetIndex, int rowIndex, User entity) {
          userList.add(entity);
        }

        @Override
        public void onError(int sheetIndex, int rowIndex, List<ExcelErrorField> errorFields) {
          errorMap.put(rowIndex, errorFields);
        }
      });

  Map<String, Object> model = new HashMap<>();
  model.put("data", userList);
  model.put("error", errorMap);
  return ResponseEntity.ok(model);
}
```

### 4. 一行代码执行 Excel 批量导出.
> 基于 Apache POI SXSSF 系列API实现导出, 大幅优化导出性能.

``` java
@RequestMapping(value = "/downXlsx", method = RequestMethod.GET)
public void downXlsx(HttpServletResponse response) {
  List<User> userList = DbUtil.getUserList(100000);
  ExcelKit.$Export(User.class, response).downXlsx(userList, false);
}
```
