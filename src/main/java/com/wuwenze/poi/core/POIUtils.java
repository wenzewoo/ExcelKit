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
package com.wuwenze.poi.core;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.OutputStream;
import java.net.URLEncoder;

/**
 * <p>
 * Apache POI SXSS相关API的简易封装
 * <p>
 * 
 * @author wuwz
 * @since 2017年4月10日
 */
public class POIUtils {
	private static final int mDefaultRowAccessWindowSize = 100;

	public static SXSSFWorkbook newSXSSFWorkbook(int rowAccessWindowSize) {
		return new SXSSFWorkbook(rowAccessWindowSize);
	}

	public static SXSSFWorkbook newSXSSFWorkbook() {
		return newSXSSFWorkbook(mDefaultRowAccessWindowSize);
	}
	
	public static SXSSFSheet newSXSSFSheet(SXSSFWorkbook wb,String sheetName) {
		return (SXSSFSheet) wb.createSheet(sheetName);
	}
	
	public static SXSSFRow newSXSSFRow(SXSSFSheet sheet,int index) {
		return (SXSSFRow) sheet.createRow(index);
	}
	
	public static SXSSFCell newSXSSFCell(SXSSFRow row, int index) {
		return (SXSSFCell) row.createCell(index);
	}

	/**
	 * 设定单元格宽度 (手动/自动)
	 * @param sheet 工作薄对象
	 * @param index 单元格索引
	 * @param width 指定宽度,-1为自适应
	 * @param value 自适应需要单元格内容进行计算
	 */
	public static void setColumnWidth(SXSSFSheet sheet,int index, short width, String value) {
		if (width == -1 && value != null && !"".equals(value)) {
			sheet.setColumnWidth(index, (short) (value.length() * 250));
		} else {
			width = width == -1 ? 200 : width;
			sheet.setColumnWidth(index, (short) (width * 35.7));
		}
	}
	
	public static void writeByLocalOrBrowser(HttpServletResponse response, String fileName, SXSSFWorkbook wb,
			OutputStream out) throws Exception {
		if(response != null) {
			// response对象不为空,响应到浏览器下载
			response.setContentType(Const.XLSX_CONTENT_TYPE);
			response.setHeader("Content-disposition", "attachment; filename=" + URLEncoder.encode(String.format("%s%s", fileName, Const.XLSX_SUFFIX), "UTF-8"));
			if (out == null) {
				out = response.getOutputStream();
			}
		}
		wb.write(out);
		out.flush();
		out.close();
	}
	/**
	 * 设置某些列的值只能输入预制的数据,显示下拉框.
	 * @param sheet 要设置的sheet.
	 * @param textlist 下拉框显示的内容
	 * @param firstRow 开始行
	 * @param endRow 结束行
	 * @param firstCol   开始列
	 * @param endCol  结束列
	 * @return 设置好的sheet.
	 */
	public static SXSSFSheet setHSSFValidation(SXSSFSheet sheet,
											  String[] textlist, int firstRow, int endRow, int firstCol,
											  int endCol) {
		DataValidationHelper validationHelper = sheet.getDataValidationHelper();
		// 加载下拉列表内容
		DataValidationConstraint explicitListConstraint = validationHelper.createExplicitListConstraint(textlist);
		// 设置数据有效性加载在哪个单元格上,四个参数分别是：起始行、终止行、起始列、终止列
		CellRangeAddressList regions = new CellRangeAddressList(firstRow,endRow, firstCol, endCol);
		// 数据有效性对象
		DataValidation validation = validationHelper.createValidation(explicitListConstraint, regions);
		validation.setSuppressDropDownArrow(true);
		validation.createErrorBox("tip","请从下拉列表选取");
		//错误警告框
		validation.setShowErrorBox(true);
		sheet.addValidationData(validation);
		return sheet;
	}

	public static void checkExcelFile(File file) {
		if(file == null || !file.exists()) {
			throw new IllegalArgumentException("excel文件不存在.");
		}
		
		checkExcelFile(file.getAbsolutePath());
	}
	public static void checkExcelFile(String file) {
		if(!file.endsWith(Const.XLSX_SUFFIX)) {
			throw new IllegalArgumentException("抱歉,目前ExcelKit仅支持.xlsx格式的文件.");
		}
	}

	
}
