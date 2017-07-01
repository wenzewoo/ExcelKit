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
package org.wuwz.poi.core;

import java.io.File;
import java.io.OutputStream;
import java.net.URLEncoder;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

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
			sheet.setColumnWidth(index, (short) (value.length() * 512));
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
