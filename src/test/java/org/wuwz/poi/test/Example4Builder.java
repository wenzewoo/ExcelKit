package org.wuwz.poi.test;

import org.wuwz.poi.ExcelKit;
import org.wuwz.poi.test.entity.User;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.List;

public class Example4Builder {

    /**
     * 生成Excel文件示例, 注意(生成的文件大小取决于JVM的可用内存)
     * @throws FileNotFoundException
     */
    public void builder(File excelFile) throws FileNotFoundException {
        List<User> users = Db.getUsers();

        // 生成excel文件到本地.
        ExcelKit.$Builder(User.class)
                .setMaxSheetRecords(10000) // 每个sheet的记录数, 默认为10000可以不设置
                .toExcel(users,"用户信息", new FileOutputStream(excelFile));
    }
}
