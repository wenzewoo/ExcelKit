package examples;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.wuwz.poi.ExcelKit;
import org.wuwz.poi.hanlder.ReadHandler;


// @WebServlet("/example")
public class ExampleServlet extends HttpServlet {

	private static final long serialVersionUID = -3872315657402568883L;

	
	/**
	 * Web端导入导出（伪代码）
	 */
	
	@Override
	protected void service(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
		
		/*1. 海量数据导入 （伪代码）*/
		
		// 获取上传文件
		File upload = null;
		
		// 读取并解析文件
		final List<Object> exportData = new ArrayList<Object>();
		ExcelKit.$Import().readExcel(upload, new ReadHandler() {
			
			@Override
			public void handler(int sheetIndex, int rowIndex, List<String> row) {
				try {
					// 排除表头
					if(rowIndex == 0) return;
					
					// 验证行数据
					if(!validRow(row)) {
						// TODO: 行数据rowIndex验证失败，记录
						
					} else {
						// 解析行数据
						
						// 方案1：记录行数据，读取完成后批量入库
						exportData.add(analysisRow(row));
						
						// 方案2：单行直接入库
						// xxx.save(analysisRow(row));
					} 
					
				} catch (Exception e) {
					e.printStackTrace();
					// 读取行：rowIndex发生异常，记录
				}
			}

			private Object analysisRow(List<String> row) {
				// TODO 解析行数据为对象
				return null;
			}

			private boolean validRow(List<String> row) {
				// TODO 验证行数据
				return true;
			}
		});
		
		// 方案1： 文件读取解析完毕, 批量入库。
		//xxx.batchSave(exportData);
		
		// TODO 响应结果，成功条数、失败条数（行、原因）
		
		
		/*2. 导出 （伪代码）*/
		List<User> users = Db.getUsers();
		ExcelKit.$Export(User.class, resp).toExcel(users, "用户数据");
	}
	

}
