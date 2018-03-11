package org.wuwz.poi.test;

import java.io.File;
import java.io.FileNotFoundException;

public class Demo {

    public static void main(String[] args) throws FileNotFoundException {
        File excelFile = new File("D:/testUserData.xlsx");

        // 生成到本地
        new Example4Builder().builder(excelFile);

        // 将本地Excel文件读取并解析入库
        new Example4Import().importExcel(excelFile);
    }
}
