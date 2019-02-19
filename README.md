# ExcelKit
> 简单、好用且轻量级的海量Excel文件导入导出解决方案。<反馈问题微信：`unknow-uid`>

## POM.xml
```xml
<dependency>
    <groupId>com.wuwenze</groupId>
    <artifactId>ExcelKit</artifactId>
    <version>2.0.71</version>
</dependency>
```

## 示例
### 1. ExcelMapping (配置Excel与实体之间的映射关系)

> 现支持两种配置方式: 注解 或者 XML
```java
@Excel("用户信息")
public class User {

  @ExcelField(value = "编号", width = 30)
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
      comment = "可以乱填，但是长度不能超过80，导入时最终会转换为数字",//
      writeConverter = CustomizeFieldWriteConverter.class,// 写文件时，将数字转字符串
      readConverter = CustomizeFieldReadConverter.class// 读文件时，将字符串转数字
  )
  private Integer customizeField;
  
  // Getter and Setter ..
}
```

> XML 配置方式, 必须需将 xml 文件放置在`classpath:excel-mapping/{entityClassName}.xml`

```java
public class User2 {

  private Integer userId;
  private String username;
  private String password;
  private String email;
  private Integer sex;
  private UserGroup userGroup;
  private Date createAt;
  private Integer customizeField;
  
  // Getter and Setter ..
}
```

> classpath:excel-mapping/com.wuwenze.entity.User2.xml
```xml
<?xml version="1.0" encoding="UTF-8"?>
<excel-mapping name="用户信息">
  <property name="userId" column="编号" width="30"/>
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
  <property name="userGroup.name" column="用户组" options="com.wuwenze.options.UserGroupNameOptions"/>
  <property name="createAt" column="创建时间" dateFormat="yyyy/MM/dd HH:mm:ss"/>
  <property
    name="customizeField"
    column="自定义字段"
    maxLength="80"
    comment="可以乱填，但是长度不能超过80，导入时最终会转换为数字"
    writeConverter="com.wuwenze.converter.CustomizeFieldWriteConverter"
    readConverter="com.wuwenze.converter.CustomizeFieldReadConverter"
  />
</excel-mapping>
```
以上两种方式2选一即可, ExcelKit 会`优先装载 XML 配置文件`.

### 2. 实现相关的转换器
> 通过实现`com.wuwenze.poi.config.Options`自定义导入模板的下拉框数据源。

```java
public class UserGroupNameOptions implements Options {

  @Override
  public String[] get() {
    return new String[]{"管理组", "普通会员组", "游客"};
  }
}
```

> 通过实现`com.wuwenze.poi.validator.Validator`自定义单元格导入时的验证规则。
```java
public class UsernameValidator implements Validator {
  final List<String> ERROR_USERNAME_LIST = Lists.newArrayList(
      "admin", "root", "master", "administrator", "sb"
  );

  @Override
  public String valid(Object o) {
    String username = (String) o;
    if (username.length() > 12) {
      return "用户名不能超过12个字符。";
    }

    if (ERROR_USERNAME_LIST.contains(username)) {
      return "用户名非法，不允许使用。";
    }
    return null;
  }
}
```

> 实现`com.wuwenze.poi.convert.WriteConverter`以及`com.wuwenze.poi.convert.ReadConverter`单元格读写转换器。
```java
public class CustomizeFieldWriteConverter implements WriteConverter {

  /**
   * 写文件时，将值进行转换（此处示例为将数值拼接为指定格式的字符串）
   */
  @Override
  public String convert(Object o) throws ExcelKitWriteConverterException {
    return (o + "_convertedValue");
  }
}

public class CustomizeFieldReadConverter implements ReadConverter {

  /**
   * 读取单元格时，将值进行转换（此处示例为计算单元格字符串char的总和）
   */
  @Override
  public Object convert(Object o) throws ExcelKitReadConverterException {
    String value = (String) o;

    int convertedValue = 0;
    for (char c : value.toCharArray()) {
      convertedValue += Integer.valueOf(c);
    }
    return convertedValue;
  }
}
```

### 3. 一行代码构建 Excel 导入模板
> 使用 ExcelKit 提供的API 构建导入模板, 会根据配置生成批注, 下拉框等
``` java
// 生成导入模板（含3条示例数据）
@RequestMapping(value = "/downTemplate", method = RequestMethod.GET)
public void downTemplate(HttpServletResponse response) {
  List<User> userList = DbUtil.getUserList(3);
  ExcelKit.$Export(User.class, response).downXlsx(userList, true);
}
```
![](https://cdn.nlark.com/yuque/0/2019/png/243237/1550539800515-e9714f25-e415-4e70-a4b9-2b5a229dfce0.png)  
![](https://cdn.nlark.com/yuque/0/2019/png/243237/1550539833934-6b7b2ca8-c7a0-4872-a1a9-722fc5c403d1.png)


### 4. 执行文件导入
> 使用边读边处理的方式, 无需担心内存溢出, 也不用理会 Excel 文件到底有多大.

``` java
@RequestMapping(value = "/importUser", method = RequestMethod.POST)
public ResponseEntity<?> importUser(@RequestParam MultipartFile file)
    throws IOException {
  long beginMillis = System.currentTimeMillis();

  List<User> successList = Lists.newArrayList();
  List<Map<String, Object>> errorList = Lists.newArrayList();

  ExcelKit.$Import(User.class)
      .readXlsx(file.getInputStream(), new ExcelReadHandler<User>() {

        @Override
        public void onSuccess(int sheetIndex, int rowIndex, User entity) {
          successList.add(entity); // 单行读取成功，加入入库队列。
        }

        @Override
        public void onError(int sheetIndex, int rowIndex,
            List<ExcelErrorField> errorFields) {
          // 读取数据失败，记录了当前行所有失败的数据
          errorList.add(MapUtil.newHashMap(//
              "sheetIndex", sheetIndex,//
              "rowIndex", rowIndex,//
              "errorFields", errorFields//
          ));
        }
      });

  // TODO: 执行successList的入库操作。

  return ResponseEntity.ok(MapUtil.newHashMap(
      "data", successList,
      "haveError", !CollectionUtil.isEmpty(errorList),
      "error", errorList,
      "timeConsuming", (System.currentTimeMillis() - beginMillis) / 1000L
  ));
}
```

> 全部导入成功示例：

![](https://cdn.nlark.com/yuque/0/2019/png/243237/1550539857806-7dd5b511-4cfe-43f7-8d82-f76bd831e7af.png)

> 部分导入失败示例（包含错误信息）：

![](https://cdn.nlark.com/yuque/0/2019/png/243237/1550539878743-a0cc24a2-1f30-4ccc-8d14-225a3bad30f5.png)


### 5. 一行代码执行 Excel 批量导出.
> 基于 Apache POI SXSSF 系列API实现导出, 大幅优化导出性能.

``` java
@RequestMapping(value = "/downXlsx", method = RequestMethod.GET)
public void downXlsx(HttpServletResponse response) {
  long beginMillis = System.currentTimeMillis();
  List<User> userList = DbUtil.getUserList(100000);// 生成10w条测试数据
  ExcelKit.$Export(User.class, response).downXlsx(userList, false);
  log.info("#ExcelKit.$Export success, size={},timeConsuming={}s",//
      userList.size(), (System.currentTimeMillis() - beginMillis) / 1000L);
}
```
![](https://cdn.nlark.com/yuque/0/2019/png/243237/1550539922063-7681b3df-6ccf-4f42-939b-269c84b8f8bc.png)  
需要注意的是，虽然ExcelKit针对导出做了大量优化，但导出数据也需要量力而行。  
![](https://cdn.nlark.com/yuque/0/2019/png/243237/1550539967120-9eff43bf-d225-4268-8685-2b441f7024d5.png)