# ExcelKit

> 简单、好用且轻量级的海量Excel文件导入导出解决方案。<反馈交流 QQ 群: 728244066>

***注意: 当前版本为全新重构的新版本, 不兼容老版本, 新版正在测试阶段, 请勿直接用于生产环境!***

## POM.xml 
> (测试版不上传中央仓库, 请手动 clone 后 install 到本地仓库使用)

```xml
<dependency>
    <groupId>com.wuwenze</groupId>
    <artifactId>ExcelKit</artifactId>
    <version>2.0.7.beta(20180519)</version>
</dependency>
```

## 示例

> 完整示例项目: https://gitee.com/wuwenze/ExcelKit-Examples

### 1. ExcelMapping

> 现支持两种配置方式: 注解 或者 XML
```java
/**
 * 注解方式使用 ExcelKit
 */
@Excel("用户信息")
public class UserForAnno {

    private Integer id;

    @ExcelField(value = "姓名", required = true)
    private String name;

    @ExcelField(value = "手机号", validator = MobileValidator.class)
    private String tel;

    @ExcelField(
            value = "性别",//
            readConverterExp = "未知=0,男=1,女=2",//
            writeConverterExp = "0=未知,1=男,2=女",//
            options = SexOptions.class//
    )
    private Integer sex;

    @ExcelField(value = "用户组", name = "userGroup.name", options = UserGroupOptions.class)
    private UserGroup userGroup;

    @ExcelField(value = "创建时间", dateFormat = "yyyy/MM/dd")
    private Date createAt;

    @ExcelField(value = "自定义字段1", comment = "提示: 只能填写数字", regularExp = "[0-9]+", regularExpMessage = "必须是数字")
    private String field1;

    @ExcelField(//
            value = "自定义字段2",//
            maxLength = 80,
            comment = "可以乱填, 但是长度不能超过80",
            writeConverter = Field2WriteConverterConverter.class,//
            readConverter = Field2ReadConverterConverter.class//
    )
    private Integer field2;
}
```

> XML 配置方式, 需将 xml 文件放置在classpath:excel-mapping/{entityClassName}.xml

> 如: classpath:excel-mapping/com.wuwenze.excelkitexamples.entity.UserForXml.xml
```java
/**
 * XML方式使用 ExcelKit
 */
public class UserForXml {
    private Integer id;
    private String name;
    private String tel;
    private Integer sex;
    private UserGroup userGroup;
    private Date createAt;
    private String field1;
    private Integer field2;
}
```

```xml
<?xml version="1.0" encoding="UTF-8"?>
<excel-mapping name="用户信息">
    <property name="name" column="姓名" required="true"/>
    <property name="tel" column="手机号" validator="com.wuwenze.poi.validator.MobileValidator"/>
    <property name="sex"
              column="性别"
              readConverterExp="未知=0,男=1,女=2"
              writeConverterExp="0=未知,1=男,2=女"
              options="com.wuwenze.excelkitexamples.options.SexOptions"
    />
    <property name="userGroup.name" column="用户组"
              options="com.wuwenze.excelkitexamples.options.UserGroupOptions"/>
    <property name="createAt" column="创建时间" dateFormat="yyyy/MM/dd"/>
    <property name="field1"
              column="自定义字段1"
              comment="提示: 只能填写数字"
              regularExp="[0-9]+"
              regularExpMessage="必须是数字"/>
    <property name="field2"
              column="自定义字段2"
              maxLength="80"
              comment="可以乱填, 但是长度不能超过80"
              writeConverter="com.wuwenze.excelkitexamples.converter.Field2WriteConverterConverter"
              readConverter="com.wuwenze.excelkitexamples.converter.Field2ReadConverterConverter"
    />
</excel-mapping>
```
以上两种方式2选一即可, ExcelKit 会优先装载 XML 配置文件.


### 2. 一行代码构建 Excel 导入模板
> 使用 ExcelKit 提供的API 构建导入模板, 会根据配置生成批注, 下拉框等
``` java
@GetMapping("/down_template_by_xml")
public void downTemplate(HttpServletResponse response) {
    // 构造演示数据
    List<UserForXml> exampleRows = Lists.newArrayList(//
            UserForXml.builder()
                    .name("姓名")
                    .sex(1)
                    .tel("17311223344")
                    .createAt(new Date())
                    .userGroup(UserGroup.builder().name("研发部").build())
                    .field1("121233")
                    .field2(0)
                    .build()
    );

    // 构建并下载模板
    ExcelKit.$Export(UserForXml.class, response).downXlsx(exampleRows, true);
}
```

### 3. 执行文件导入
> 使用边读边处理的方式, 无需担心内存溢出, 也不用理会 Excel 文件到底有多大.

``` java
@PostMapping("/import")
public ResponseEntity<?> importExcel(@RequestParam("excelFile") MultipartFile file) throws IOException {
    // 省略文件检查..

    // 执行文件导入.
    long beginTimeMillis = System.currentTimeMillis();
    final List<UserForAnno> data = Lists.newArrayList();
    final List<Map<String, Object>> error = Lists.newArrayList();
    ExcelKit.$Import(UserForAnno.class)//
            .readXlsx(file.getInputStream(), new ExcelReadHandler<UserForAnno>() {
                @Override
                public void onSuccess(int sheet, int row, UserForAnno userForAnno) {
                    // 当前行读取成功, 入库或加入批量入库队列.
                    data.add(userForAnno);
                }

                @Override
                public void onError(int sheet, int row, List<ExcelErrorField> errorFields) {
                    // 当前行读取失败, 获取失败详情.
                    error.add(ImmutableMap.of("row", row, "errorFields", errorFields));
                }
            });
    long time = ((System.currentTimeMillis() - beginTimeMillis) / 1000L);
    System.out.println("数据量: " + data.size() + ", 耗时: " + time + "秒");
    ImmutableMap<String, Object> retJsonMap = ImmutableMap.of(//
            "time", "耗时" + time + "秒",
            "data", data,
            "error", error
    );
    return ResponseEntity.ok(retJsonMap);
}
```

### 4. 一行代码执行 Excel 批量导出.
> 基于 Apache POI SXSSF 系列API实现导出, 大幅优化导出性能.

``` java
@GetMapping("/export")
public void export(HttpServletResponse response) {
    List<UserForAnno> rows = Lists.newArrayList();
    for (int i = 0; i < 100000; i++) {
        UserForAnno build = UserForAnno.builder()
                .name("name." + i)
                .sex(i % 2 == 0 ? 1 : 2)
                .tel("17311223344")
                .createAt(new Date())
                .userGroup(UserGroup.builder().name("研发部").build())
                .field1("121233")
                .field2(0)
                .build();
        rows.add(build);
    }
    long beginTimeMillis = System.currentTimeMillis();
    ExcelKit.$Export(UserForAnno.class, response).downXlsx(rows, false);
    long time = ((System.currentTimeMillis() - beginTimeMillis) / 1000L);
    System.out.println("数据量: " + rows.size() + ", 耗时: " + time + "秒");
}
```

***有关新版本的详细 API, 我将抽空完善; 另外当前版本可能 BUG 满天飞, 无需惊讶, 请果断提Issues.***