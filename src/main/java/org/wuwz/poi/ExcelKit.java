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

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.codec.binary.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.wuwz.poi.annotation.ExportConfig;
import org.wuwz.poi.convert.ExportConvert;
import org.wuwz.poi.core.POIUtils;
import org.wuwz.poi.core.XlsxReader;
import org.wuwz.poi.hanlder.ExportHandler;
import org.wuwz.poi.hanlder.ReadHandler;
import org.wuwz.poi.pojo.ExportItem;

/**
 * <p>
 * ExcelKit.java : Excel快速导入导出工具 <br/>
 * 简单、好用且轻量级的海量Excel文件导入导出解决方案. (仅提供excel2007以及更高版本的支持)
 * <p>
 * 
 * @author wuwz
 * @since 2017年4月10日
 */
public class ExcelKit {
	private static Logger log = Logger.getLogger(ExcelKit.class);

	private Class<?> mClass = null;
	private HttpServletResponse mResponse = null;
	// 默认以此值填充空单元格,可通过 setEmptyCellValue(string)改变其默认值。
	private String mEmptyCellValue = null;
	// 分Sheet机制：每个Sheet最多多少条数据
	private Integer mMaxSheetRecords = 10000;
	// 缓存数据格式器实例,避免多次使用反射进行实例化
	private Map<String, ExportConvert> mConvertInstanceCache = new HashMap<String, ExportConvert>();

	protected ExcelKit() {}

	protected ExcelKit(Class<?> clazz) {
		this(clazz, null);
	}

	protected ExcelKit(Class<?> clazz, HttpServletResponse response) {
		this.mResponse = response;
		this.mClass = clazz;
	}

	/**
	 * 用于生成本地文件
	 * 
	 * @param clazz
	 *            实体Class对象
	 * @return ExcelKit
	 */
	public static ExcelKit $Builder(Class<?> clazz) {
		return new ExcelKit(clazz);
	}

	/**
	 * 用于浏览器导出
	 * 
	 * @param clazz
	 *            实体Class对象
	 * @param response
	 *            原生HttpServletResponse对象
	 * @return ExcelKit
	 */
	public static ExcelKit $Export(Class<?> clazz, HttpServletResponse response) {
		return new ExcelKit(clazz, response);
	}

	/**
	 * 用于导入数据解析
	 * 
	 * @return ExcelKit
	 */
	public static ExcelKit $Import() {
		return new ExcelKit();
	}

	/**
	 * 读取Excel时以该值填充空单元格值 (默认null)
	 * 
	 * @param emptyCellValue
	 *            单元格值
	 * @return this
	 */
	public ExcelKit setEmptyCellValue(String emptyCellValue) {
		this.mEmptyCellValue = emptyCellValue;
		return this;
	}

	/**
	 * 分Sheet机制：每个Sheet最多多少条数据(默认10000)
	 * 
	 * @param size
	 *            数据条数
	 * @return this
	 */
	public ExcelKit setMaxSheetRecords(Integer size) {
		this.mMaxSheetRecords = size;
		return this;
	}

	/**
	 * 导出Excel(此方式需依赖浏览器实现文件下载,故应先使用$Export()构造器)
	 * 
	 * @param data
	 *            数据集合
	 * @param sheetName
	 *            工作表名字
	 * @return true-操作成功,false-操作失败
	 */
	public boolean toExcel(List<?> data, String sheetName) {
		required$ExportParams();

		try {
			return toExcel(data, sheetName, mResponse.getOutputStream());
		} catch (IOException e) {
			log.error("导出Excel失败:" + e.getMessage(), e);
		}
		return false;
	}

	/**
	 * 针对转换方法的默认实现(提供默认样式和文件命名规则)
	 * 
	 * @param data
	 *            数据集合
	 * @param sheetName
	 *            工作表名字
	 * @param out
	 *            输出流
	 * @return true-操作成功,false-操作失败
	 */
	public boolean toExcel(List<?> data, String sheetName, OutputStream out) {

		return toExcel(data, sheetName, new ExportHandler() {

			@Override
			public CellStyle headCellStyle(SXSSFWorkbook wb) {
				CellStyle cellStyle = wb.createCellStyle();
				Font font = wb.createFont();
				cellStyle.setFillForegroundColor((short) 12);
				cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);// 填充模式
				cellStyle.setBorderTop(CellStyle.BORDER_THIN);// 上边框为细边框
				cellStyle.setBorderRight(CellStyle.BORDER_THIN);// 右边框为细边框
				cellStyle.setBorderBottom(CellStyle.BORDER_THIN);// 下边框为细边框
				cellStyle.setBorderLeft(CellStyle.BORDER_THIN);// 左边框为细边框
				cellStyle.setAlignment(CellStyle.ALIGN_LEFT);// 对齐
				cellStyle.setFillForegroundColor(HSSFColor.GREEN.index);
				cellStyle.setFillBackgroundColor(HSSFColor.GREEN.index);
				font.setBoldweight(Font.BOLDWEIGHT_NORMAL);
				// font.setFontHeightInPoints((short) 12);// 字体大小
				font.setColor(HSSFColor.WHITE.index);
				// 应用标题字体到标题样式
				cellStyle.setFont(font);
				return cellStyle;
			}

			@Override
			public String exportFileName(String sheetName) {
				return String.format("导出-%s-%s", sheetName, System.currentTimeMillis());
			}
		}, out);
	}

	public boolean toExcel(List<?> data, String sheetName, ExportHandler handler, OutputStream out) {
		required$BuilderParams();
		long begin = System.currentTimeMillis();

		if (data == null || data.size() < 1) {
			log.error("没有检测到数据,不执行导出操作。");
			return false;
		}
		
		log.info(String.format("即将导出excel数据：%s条,请稍后..", data.size()));

		// 导出列查询。
		ExportConfig currentExportConfig = null;
		ExportItem currentExportItem = null;
		List<ExportItem> exportItems = new ArrayList<ExportItem>();
		for (Field field : mClass.getDeclaredFields()) {

			currentExportConfig = field.getAnnotation(ExportConfig.class);
			if (currentExportConfig != null) {
				currentExportItem = new ExportItem()
						.setField(field.getName())
						.setDisplay("field".equals(currentExportConfig.value()) ? field.getName() : currentExportConfig.value())
						.setWidth(currentExportConfig.width())
						.setConvert(currentExportConfig.convert())
						.setColor(currentExportConfig.color())
						.setReplace(currentExportConfig.replace());
				exportItems.add(currentExportItem);
			}
			
			currentExportItem = null;
			currentExportConfig = null;
		}

		// 创建新的工作薄。
		SXSSFWorkbook wb = POIUtils.newSXSSFWorkbook();

		double sheetNo = Math.ceil(data.size() / mMaxSheetRecords);// 取出一共有多少个sheet.

		// =====多sheet生成填充数据=====
		for (int index = 0; index <= (sheetNo == 0.0 ? sheetNo : sheetNo - 1); index++) {
			SXSSFSheet sheet = POIUtils.newSXSSFSheet(wb, sheetName + (index == 0 ? "" : "_" + index));

			// 创建表头
			SXSSFRow headerRow = POIUtils.newSXSSFRow(sheet, 0);
			for (int i = 0; i < exportItems.size(); i++) {
				SXSSFCell cell = POIUtils.newSXSSFCell(headerRow, i);
				POIUtils.setColumnWidth(sheet, i, exportItems.get(i).getWidth(), exportItems.get(i).getDisplay());
				cell.setCellValue(exportItems.get(i).getDisplay());
				
				CellStyle style = handler.headCellStyle(wb);
				if (style != null) {
					cell.setCellStyle(style);
				}
			}
			
			// 产生数据行
			if (data != null && data.size() > 0) {

				int startNo = index * mMaxSheetRecords;
				int endNo = Math.min(startNo + mMaxSheetRecords, data.size());
				for (int i = startNo; i < endNo; i++) {
					SXSSFRow bodyRow = POIUtils.newSXSSFRow(sheet, i + 1 - startNo);

					for (int j = 0; j < exportItems.size(); j++) {
						// 处理单元格值
						String cellValue = exportItems.get(j).getReplace();
						if ("".equals(cellValue)) {
							try {
								cellValue = BeanUtils.getProperty(data.get(i), exportItems.get(j).getField());
							} catch (Exception e) {
								log.error("获取" + exportItems.get(j).getField() + "的值失败.", e);
							}
						}

						// 格式化单元格值
						if (!"".equals(exportItems.get(j).getConvert())) {
							cellValue = convertCellValue(Integer.parseInt(cellValue), exportItems.get(j).getConvert());
						}

						// 单元格宽度
						POIUtils.setColumnWidth(sheet, j, exportItems.get(j).getWidth(), cellValue);

						SXSSFCell cell = POIUtils.newSXSSFCell(bodyRow, j);
						// fix: 当值为“”时,当前index的cell会失效
						cell.setCellValue("".equals(cellValue) ? null : cellValue);
						
						CellStyle style = wb.createCellStyle();
						Font font = wb.createFont();
						font.setColor(exportItems.get(j).getColor());
						style.setFont(font);
						cell.setCellStyle(style);
					}
				}
			}
		}

		try {
			// 生成Excel文件并下载.(通过response对象是否为空来判定是使用浏览器下载还是直接写入到output中)
			POIUtils.writeByLocalOrBrowser(mResponse, handler.exportFileName(sheetName), wb, out);
		} catch (Exception e) {
			log.error("生成Excel文件失败:" + e.getMessage(), e);
			return false;
		}

		log.info(String.format("Excel处理完成,共生成数据:%s行 (不包含表头),耗时：%s seconds.", (data != null ? data.size() : 0),
				(System.currentTimeMillis() - begin) / 1000F));
		return true;
	}

	/**
	 * 读取Excel数据(使用SAX的方式进行解析,读取所有Sheet数据)
	 * 
	 * @param excelFile
	 *            excel文件
	 * @param handler
	 *            数据处理回调
	 */
	public void readExcel(File excelFile, ReadHandler handler) {
		this.readExcel(excelFile, -1, handler);
	}

	/**
	 * 读取Excel（使用SAX的方式进行解析,读取指定Sheet数据）
	 * 
	 * @param excelFile
	 *            excel文件
	 * @param sheetIndex
	 *            sheet索引,-1为读取所有
	 * @param handler
	 *            数据处理回调
	 */
	public void readExcel(File excelFile, int sheetIndex, ReadHandler handler) {
		long begin = System.currentTimeMillis();
		String fileName = excelFile.getAbsolutePath();
		XlsxReader reader = new XlsxReader(handler).setEmptyCellValue(mEmptyCellValue);
		try {
			if (sheetIndex >= 0) {
				// 读取指定sheet
				reader.process(fileName, sheetIndex);
			} else {
				// 读取所有sheet
				reader.process(fileName);
			}
		} catch (Exception e) {
			log.error("读取excel文件时发生异常：" + e.getMessage(), e);
		}
		
		log.info(String.format("Excel读取并处理完成,耗时：%s seconds.",(System.currentTimeMillis() - begin) / 1000F));
	}

	// convertCellValue: number to String (beta)
	private String convertCellValue(Integer oldValue, String format) {
		try {
			String protocol = format.split(":")[0];

			// 键值对字符串解析：s:1=男,2=女
			if ("s".equalsIgnoreCase(protocol)) {

				String[] pattern = format.split(":")[1].split(",");
				for (String p : pattern) {
					String[] cp = p.split("=");

					if (Integer.parseInt(cp[0]) == oldValue) {
						return cp[1];
					}
				}

			}
			// 使用处理类进行处理：c:org.wuwz.poi.test.GradeCellFormat
			if ("c".equalsIgnoreCase(protocol)) {
				String clazz = format.split(":")[1];
				ExportConvert export = mConvertInstanceCache.get(clazz);

				if (export == null) {
					export = (ExportConvert) Class.forName(clazz).newInstance();
					mConvertInstanceCache.put(clazz, export);
				}

				if (mConvertInstanceCache.size() > 10)
					mConvertInstanceCache.clear();

				return export.handler(oldValue);
			}
		} catch (Exception e) {
			log.error("出现问题,可能是@ExportConfig.format()的值不规范导致。", e);
		}
		return String.valueOf(oldValue);
	}

	private void required$BuilderParams() {
		if (mClass == null) {
			throw new IllegalArgumentException("请先使用org.wuwz.poi.ExcelKit.$Builder(Class<?>)构造器初始化参数。");
		}
	}

	private void required$ExportParams() {
		if (mClass == null || mResponse == null) {
			throw new IllegalArgumentException(
					"请先使用org.wuwz.poi.ExcelKit.$Export(Class<?>, HttpServletResponse)构造器初始化参数。");
		}

	}

}
