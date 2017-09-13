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

import org.wuwz.poi.convert.ExportConvert;

import java.util.HashMap;
import java.util.Map;

public class GradeIdConvert implements ExportConvert {
	
	static Map<Integer, String> records;
	static {
		//默认数据库字典查询 select * from tb_grades
		records = new HashMap<Integer, String>();
		records.put(1, "一年级学生");
		records.put(2, "二年级学生");
		records.put(3, "三年级学生");
	}

	@Override
	public String handler(Integer val) {
		return records.get(val) != null ? records.get(val) : "没有该记录";
	}

}
