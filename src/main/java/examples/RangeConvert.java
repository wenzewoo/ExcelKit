package examples;

import org.wuwz.poi.convert.ExportRange;

/**
 * 下拉框定义
 * @author wuwz
 */
public class RangeConvert implements ExportRange {
	private static String[] demo = null;

	static {
		demo = new String[] { "列表1", "列表2" };
	}

	@Override
	public String[] handler() {
		return demo;
	}
}
