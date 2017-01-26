package org.wuwz.poi;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel导出项配置
 * @author wuwz
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.FIELD })
public @interface ExportConfig {

	
	/**
	 * 表头显示名,默认为字段名
	 */
	String value() default "field";
	
	/**
	 * 宽度
	 */
	short width() default 300;
	
	/**
	 * 是否导出数据(如果不导出数据,Ĭ将以“******”填充单元格)
	 */
	boolean isExportData() default true;
}
