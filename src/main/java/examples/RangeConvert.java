package examples;

import org.wuwz.poi.convert.ExportRange;

public class RangeConvert implements ExportRange {
    private static String[] demo = null;

    static {
        demo =new String[]{"列表1", "列表2"};
    }

    @Override
    public String[] handler() {
        return demo;
    }
}
