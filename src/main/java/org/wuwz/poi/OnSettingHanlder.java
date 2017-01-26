package org.wuwz.poi;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 导出Excel设置接口。
 * @author wuwz
 */
public interface OnSettingHanlder {
	/**
	 * 设置表头样式
	 */
	CellStyle getHeadCellStyle(Workbook wb);

	/**
	 * 设置单元格样式
	 */
	CellStyle getBodyCellStyle(Workbook wb);
	
	/**
	 * 设置导出的文件名（无需处理后缀）
	 */
	String getExportFileName(String sheetName);
}
