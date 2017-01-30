package org.wuwz.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.Lists;

/**
 * Excel导入导出工具 <br>
 * Maven依赖:
 * <pre>
	需要手动添加的依赖包:
	&lt;dependency>
		&lt;groupId>javax&lt;/groupId>
		&lt;artifactId>javaee-api&lt;/artifactId>
		&lt;version>7.0&lt;/version>
	&lt;/dependency>
	&lt;dependency>
		&lt;groupId>javax.servlet&lt;/groupId>
		&lt;artifactId>javax.servlet-api&lt;/artifactId>
		&lt;version>3.1.0&lt;/version>
	&lt;/dependency>
	
 	已经集成的依赖包:
	&lt;dependency>
		&lt;groupId>org.apache.poi&lt;/groupId>
		&lt;artifactId>poi-ooxml&lt;/artifactId>
		&lt;version>3.8&lt;/version>
	&lt;/dependency>
	
	&lt;dependency>
	    &lt;groupId>com.google.guava&lt;/groupId>
	    &lt;artifactId>guava&lt;/artifactId>
	    &lt;version>18.0&lt;/version>
	&lt;/dependency>

	&lt;dependency>
	    &lt;groupId>commons-beanutils&lt;/groupId>
	    &lt;artifactId>commons-beanutils&lt;/artifactId>
	    &lt;version>1.9.3&lt;/version>
	&lt;/dependency>
 * </pre>
 * @author wuwz
 */
public class ExcelKit {
	private static Logger log = Logger.getLogger(ExcelKit.class);
	
	private Class<?> _class;
	public HttpServletResponse _response;
	/**默认以此值填充空单元格**/
	public String _emptyCellValue = "EMPTY_CELL_VALUE";
	
	private ExcelKit() {}
	private ExcelKit(Class<?> _class) {
		this(_class, null);
	}
	private ExcelKit(Class<?> _class,HttpServletResponse response) {
		this._response = response;
		this._class = _class;
	}
	
	/**用于生成本地文件**/
	public static ExcelKit $Builder(Class<?> _class) {
		return new ExcelKit(_class);
	}
	
	/**用于浏览器导出**/
	public static ExcelKit $Export(Class<?> _class,HttpServletResponse response) {
		return new ExcelKit(_class,response);
	}
	
	/**用于导入数据解析**/
	public static ExcelKit $Import() {
		return new ExcelKit();
	}
	
	public ExcelKit setEmptyCellValue(String emptyCellValue) {
		this._emptyCellValue = emptyCellValue;
		return this;
	}
	
	/**
	 * 导出Excel
	 */
	public boolean toExcel(List<?> data,String sheetName) {
		if(_response == null) {
			throw new RuntimeException("请先初始化  HttpServletResponse 对象! (通过调用 $Export(Class<?> _class,HttpServletResponse response))");
		}
		try {
			return toExcel(data, sheetName, _response.getOutputStream());
		} catch (IOException e) {
			e.printStackTrace();
		}
		return false;
	}

	public boolean toExcel(List<?> data, String sheetName, OutputStream out) {

		return toExcel(data, sheetName, ExcelType.EXCEL2007, new OnSettingHanlder() {

			@Override
			public CellStyle getHeadCellStyle(Workbook wb) {
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
				font.setFontHeightInPoints((short) 14);// 字体大小
				font.setColor(HSSFColor.WHITE.index);
				// 应用标题字体到标题样式
				cellStyle.setFont(font);
				return cellStyle;
			}

			@Override
			public CellStyle getBodyCellStyle(Workbook wb) {
				return null;
			}
			
			@Override
			public String getExportFileName(String sheetName) {
				return String.format("导出-%s-%s", sheetName,System.currentTimeMillis());
			}
		}, out);
	}

	public boolean toExcel(List<?> data, String sheetName, ExcelType type, OnSettingHanlder handler,
			OutputStream out) {
		long begin = System.currentTimeMillis();
		
		if (data == null || data.size() < 1) {
			if(log.isDebugEnabled())
				log.debug("没有检测到导出数据，将生成 Excel导入模版。");
		}
		
		// 导出列查询。
		List<ExportItem> exportItems = Lists.newArrayList();
		for (Field field : _class.getDeclaredFields()) {

			ExportConfig ec = field.getAnnotation(ExportConfig.class);
			if (ec != null) {
				exportItems.add(
					new ExportItem.Builder()
						.setField(field.getName())
						.setDisplay("field".equals(ec.value()) ? field.getName() : ec.value())
						.setIsExportData(ec.isExportData())
						.setWidth(ec.width())
						.create()
				);
			}
		}

		// 创建工作薄。
		Workbook wb = this.createWorkbook(type);

		// 创建Sheet
		Sheet sheet = wb.createSheet(sheetName);

		// 创建表头
		Row header = sheet.createRow(0);
		for (int i = 0; i < exportItems.size(); i++) {
			ExportItem exportItem = exportItems.get(i);
			Cell cell = header.createCell(i);
			sheet.setColumnWidth(i, (short) (exportItem.getWidth() * 35.7));
			cell.setCellValue(exportItem.getDisplay());
			CellStyle style = handler.getHeadCellStyle(wb);
			if (style != null) {
				cell.setCellStyle(style);
			}
		}

		// 产生数据行
		if(data != null && data.size() > 0) {
			
			for (int i = 0; i < data.size(); i++) {

				Row body = sheet.createRow(i + 1);

				for (int j = 0; j < exportItems.size(); j++) {
					ExportItem exportItem = exportItems.get(j);
					sheet.setColumnWidth(j, (short) (exportItem.getWidth() * 35.7));
					Cell cell = body.createCell(j);
					if (exportItem.getIsExportData()) {
						try {
							cell.setCellValue(BeanUtils.getProperty(data.get(i), exportItem.getField()));
						} catch (Exception e) {
							e.printStackTrace();
						}
					} else {
						cell.setCellValue("******");
					}

					CellStyle style = handler.getBodyCellStyle(wb);
					if (style != null) {
						cell.setCellStyle(style);
					}
				}
			}
		}

		try {
			// 生成Excel文件并下载.
			if(_response != null) {
				// 浏览器响应头
				String fileName = String.format("%s%s", handler.getExportFileName(sheetName),getExcelSuffix(type));
				_response.setContentType(getContentType(type));
				_response.setHeader("Content-disposition", "attachment; filename=" + URLEncoder.encode(fileName, "UTF-8"));
				
				if(out == null) {
					out = _response.getOutputStream();
				}
			}
			wb.write(out);
			out.flush();
			out.close();
			
			log.info(String.format("Excel处理完成,共生成数据:%s行 (不包含表头),耗时：%s seconds.", 
					(data != null ? data.size() : 0),
					(System.currentTimeMillis() - begin) / 1000F));
		} catch (IOException e) {
			e.printStackTrace();
		}

		return true;
	}
	
	/**
	 * 读取Excel数据(默认读取第一个工作表的所有数据,排除表头)
	 * @param excelFile
	 * @return
	 */
	public void readExcel(File excelFile,OnReadDataHandler handler) {
		readExcel(excelFile, handler, 0, 1, -1, 0, -1);
	}
	
	/**
	 * 读取Excel数据
	 * @param excelFile
	 * @param sheetIndex 工作表索引
	 * @param startRowIndex 开始行索引
	 * @param endRowIndex 结束行索引,-1为所有
	 * @param startCellIndex 开始列索引
	 * @param endCellIndex 结束列索引,-1为所有
	 * @return
	 */
	public void readExcel(File excelFile,OnReadDataHandler handler,int sheetIndex,int startRowIndex,int endRowIndex,int startCellIndex, int endCellIndex) {
		long totalRows = 0l;
		long begin = System.currentTimeMillis();
		Workbook wb = createWorkbook(excelFile);
		
		if(wb != null) {
			Sheet sheet = wb.getSheetAt(sheetIndex);
			
			if(sheet != null) {
				
				if(endRowIndex == -1) {
					endRowIndex = sheet.getPhysicalNumberOfRows();
				}
				if(endCellIndex == -1) {
					endCellIndex = sheet.getRow(0).getPhysicalNumberOfCells();
				}
				
				for (int i = startRowIndex; i < endRowIndex; i++) {
					List<String> rowData = Lists.newArrayList();
					Row row = sheet.getRow(i);
					if(row != null) {
						
						for (int j = startCellIndex; j < endCellIndex; j++) {
							Cell cell = row.getCell(j);
							
							rowData.add(cell != null ? cell.getStringCellValue() : _emptyCellValue);
						}
					}
					if(rowData.size() > 0) {
						// 处理当前行数据
						handler.handler(rowData); 
						totalRows++;
					}
				}
			}
		}
		log.info(String.format("Excel数据读取并处理完成,共读取数据：%s行,耗时：%s seconds.", totalRows,(System.currentTimeMillis() - begin) / 1000F));
	}

	public Workbook createWorkbook(File file) {
		Workbook workbook = null;
		try {
			workbook = new HSSFWorkbook(new FileInputStream(file));//2003
		} catch (Exception e) {
			try {
				workbook = new XSSFWorkbook(new FileInputStream(file));//2003以上
			} catch (Exception e1) {
				throw new RuntimeException("不能读取有效的Excel数据！");
			}
		}
		return workbook;
	}

	public Workbook createWorkbook(ExcelType type) {
		if (type == ExcelType.EXCEL2003)
			return new HSSFWorkbook();
		return new XSSFWorkbook();
	}
	
	public String getExcelSuffix(ExcelType type) {
		if (type == ExcelType.EXCEL2003) {
			return ".xls";
		} else {
			return ".xlsx";
		}
	}
	
	public String getContentType(ExcelType type) {
		if (type == ExcelType.EXCEL2003) {
			return "application/vnd.ms-excel";
		} else {
			return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
		}
	}
}

class ExportItem {

	private String field; // 属性名
	private String display; // 显示名
	private Short width; // 宽度
	private Boolean isExportData; // 是否导出数据

	public String getField() {
		return field;
	}

	public void setField(String field) {
		this.field = field;
	}

	public String getDisplay() {
		return display;
	}

	public void setDisplay(String display) {
		this.display = display;
	}

	public Short getWidth() {
		return width;
	}

	public void setWidth(Short width) {
		this.width = width;
	}

	public Boolean getIsExportData() {
		return isExportData;
	}

	public void setIsExportData(Boolean isExportData) {
		this.isExportData = isExportData;
	}

	public ExportItem() {
		super();
	}

	public ExportItem(Builder b) {
		this.field = b.field;
		this.display = b.display;
		this.width = b.width;
		this.isExportData = b.isExportData;
	}

	public static class Builder {
		private String field;
		private String display;
		private Short width;
		private Boolean isExportData;

		public Builder setField(String field) {
			this.field = field;
			return this;
		}

		public Builder setDisplay(String display) {
			this.display = display;
			return this;
		}

		public Builder setWidth(Short width) {
			this.width = width;
			return this;
		}

		public Builder setIsExportData(Boolean isExportData) {
			this.isExportData = isExportData;
			return this;
		}

		public ExportItem create() {
			return new ExportItem(this);
		}

	}

}
