package com.wuwenze.poi.test.convert;

import com.wuwenze.poi.convert.ExportRange;

/**
 * 生成下拉框数据
 */
public class RangeConvert implements ExportRange {

    @Override
    public String[] handler() {
        String[] testRangeData = {"下拉框1","下拉框2","下拉框4"};
        return testRangeData;
    }
}
