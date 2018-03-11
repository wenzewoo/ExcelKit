package org.wuwz.poi.test.convert;

import org.wuwz.poi.convert.ExportConvert;
import org.wuwz.poi.test.Db;

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
