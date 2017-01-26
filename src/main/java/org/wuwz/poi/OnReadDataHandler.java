package org.wuwz.poi;

import java.util.List;

/**
 * Excel数据读取回调
 * @author wuwz
 */
public interface OnReadDataHandler {

	/**
	 * 处理当前行数据
	 */
	void handler(List<String> rowData);
}
