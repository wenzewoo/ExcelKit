package com.wuwenze.poi.test;

import com.wuwenze.poi.ExcelKit;
import com.wuwenze.poi.hanlder.ReadHandler;
import com.wuwenze.poi.test.entity.User;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.PrintWriter;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.List;

public class Example4Import {

    /**
     * 导入Excel文件示例 (理论上支持无限大的Excel文件)
     *
     * @param excelFile
     */
    public void importExcel(File excelFile) {
        // 读取到的数据
        final List<User> importData = new ArrayList<User>();
        // 验证失败的错误消息
        final List<String> validErrorMessages = new ArrayList<String>();

        ExcelKit.$Import().readExcel(excelFile, new ReadHandler() {
            @Override
            public void handler(int sheetIndex, int rowIndex, List<String> row) {
                // 跳过标题列
                if (rowIndex == 0) return;

                // 跳过空行, 类似 null,null,null,null,null
                if (ExcelKit.isNullRowValue(row)) return;

                // 验证行数据是否符合规范
                if (validRow(sheetIndex, rowIndex, row)) {
                    // 解析数据
                    importData.add(analyticalData(row));
                }
            }

            private User analyticalData(List<String> row) {
                return new User()
                        .setUid(Integer.parseInt(row.get(0)))
                        .setUsername(row.get(1))
                        .setPassword(row.get(2))
                        .setSex("男".equals(row.get(3)) ? 1 : 2)
                        .setGradeId(Db.getGradeIdByName(row.get(4)))
                        .setGendex(row.get(5));
            }

            private boolean validRow(int sheetIndex, int rowIndex, List<String> row) {
                /**
                 * 验证规则:
                 * 1. 单元格数据长度是否一致
                 * ...
                 */
                boolean valid = row.size() == 6;
                if (!valid) {
                    // 记录解析失败数据
                    validErrorMessages.add(MessageFormat.format(
                            "sheetIndex:{0},rowIndex:{1}数据解析失败,规定长度为6,实际解析长度为:{2}", sheetIndex, rowIndex, row.size()));
                }
                return valid;
            }
        });

        // 执行批量入库
        batchSave(importData);

        // 总结
        Integer sum = importData.size() + validErrorMessages.size();
        System.out.println(MessageFormat.format("共有数据{0}条,其中成功{1}条,失败{2}条", sum, importData.size(), validErrorMessages.size()));

        if (!validErrorMessages.isEmpty()) {
            System.out.println("失败信息汇总:");
            for (String validErrorMessage : validErrorMessages) {
                System.out.println(validErrorMessage);
            }
        }
    }

    private void batchSave(List<User> importData) {
        String filename = "D://batchSaveResult.txt";
        // 数据的批量入库, 由于这里没有数据库环境, 直接将结果写入本地磁盘
        StringBuilder stringBuilder = new StringBuilder();
        for (User user : importData) {
            stringBuilder.append(user.toString()).append("\n");
        }
        PrintWriter outputStream = null;
        try {
            outputStream = new PrintWriter(new FileOutputStream(filename));
            outputStream.println(stringBuilder.toString());
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            if (outputStream != null) {
                outputStream.close();
            }
        }
        System.out.println("batchSave result:" + filename);
    }
}
