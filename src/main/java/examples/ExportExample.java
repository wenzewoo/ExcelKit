package examples;

import org.wuwz.poi.ExcelKit;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

/**
 * @author wuwz
 * @date 2017-12-28
 */
public class ExportExample {

    public static void main(String[] args) throws FileNotFoundException {
        /**
         * 导出本地文件测试
         */

        FileOutputStream outputStream = new FileOutputStream(new File("D:/rtest.xlsx"));

        ExcelKit.$Builder(User.class).toExcel(Db.getUsers(),"Users",outputStream);
    }
}
