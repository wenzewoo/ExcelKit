package com.wuwenze.poi.test.convert;

import com.wuwenze.poi.convert.ExportConvert;
import com.wuwenze.poi.test.Db;

/**
 * 单元格值转换器
 */
public class GradeIdConvert implements ExportConvert {

    @Override
    public String handler(Object val) {
        Integer gradeId = Integer.parseInt(val.toString());

        String gradeName = Db.getGrades().get(gradeId);
        return gradeName != null ? gradeName : "无记录";
    }
}
