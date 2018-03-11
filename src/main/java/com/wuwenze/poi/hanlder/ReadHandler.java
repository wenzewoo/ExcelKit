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
package com.wuwenze.poi.hanlder;

import java.util.List;

/**
 * <p>
 * 行级别数据读取处理回调
 * <p>
 * @author wuwz
 * @since 2017年4月10日
 */
public interface ReadHandler {

	/**
	 * 处理当前行数据
	 * @param sheetIndex 从0开始
	 * @param rowIndex 从0开始
	 * @param row 当前行数据
	 */
	void handler(int sheetIndex, int rowIndex, List<String> row);
}
