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
package org.wuwz.poi;

import java.util.List;

/**
 * Excel数据读取回调
 * @author wuwz
 */
public interface OnReadDataHandler {

	/**
	 * 处理当前行数据
	 * @param rowData 当前行数据,以rowData.get(cellIndex)的方式获取,如果cell的值为ExceklKit._emptyCellValue,则表示该单元格为空。
	 */
	void handler(List<String> rowData);
}
