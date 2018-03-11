package com.wuwenze.poi.test;

public class Example4Export {

    // 浏览器导出示例, 这部分代码需要Servlet以及WEB容器的支持
    // 以下代码演示在SpringMVC环境下如何导出数据

    /*@GetMapping("/export")
    public void export(HttpServletResponse response) {
        List<User> users = Db.getUsers();

        ExcelKit.$Export(User.class, response).toExcel(users, "用户信息");
    }*/
}
