/**

 * Copyright (c) 2017, 吴汶泽 (wuwz@live.com).

 *

 * Licensed under the Apache License, Version 2.0 (the "License");

 * you may not use this file except in compliance with the License.

 * You may obtain a copy of the License at

 *

 *      http://www.apache.org/licenses/LICENSE-2.0

 *

 * Unless required by applicable law or agreed to in writing, software

 * distributed under the License is distributed on an "AS IS" BASIS,

 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.

 * See the License for the specific language governing permissions and

 * limitations under the License.

 */
package examples;

import org.wuwz.poi.ExcelKit;
import org.wuwz.poi.hanlder.ReadHandler;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

public class TestExcelKit {
	static List<User> records;
	static {
		records = new ArrayList<User>();
		
		// 模拟2w条数据
		Random random = new Random();
		
		User user = null;
		for (int i = 0; i < 10 * 2; i++) {
			user = new User()
				.setUid(i + 1)
				.setUsername("0000"+(i+1))
				.setPassword("password")
				.setSex(random.nextInt(2) % 2 == 0 ? 1 : 2)
				.setGradeId(random.nextInt(4))
			    .setGendex("Aa")
			    .setDate(new Date())
				;
			
			records.add(user);
			user = null;
		}
		
		System.out.println("数据加载完毕。");
	}
	
	
	public static void main(String[] args) throws Exception {
		/*简单的测试案例*/
		
//		1. 导出excel: 测试本地环境的excel文件生成
//		ExcelKit.$Builder(User.class)
//			.setMaxSheetRecords(10000) //设置每个sheet的最大记录数,默认为10000,可不设置
//			.toExcel(records, "用户数据", new FileOutputStream(new File("E:/2.xlsx")));
		
		//2. 导出Excel并使用浏览器下载：需要web容器环境(tomcat)支持
//		ExcelKit.$ExportRange(User.class, response).toExcel(records, "用户数据");
		
		//3. 海量数据的读取（边读边处理）
		ExcelKit.$Import()
			.setEmptyCellValue("sss") // 设置空单元格的值,默认为null,可不设置
			.readExcel(new File("E:/2.xlsx"), new ReadHandler() {

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
//		
//		// 读取指定sheet
//		ExcelKit.$Import().readExcel(new File("c:/bigexcel.xlsx"), 1, new ReadHandler() {
//			
//			@Override
//			public void handler(int sheetIndex, int rowIndex, List<String> row) {
//				
//			}
//		});
	}

}
