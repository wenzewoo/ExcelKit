/**

 * Copyright (c) 2017, ������ (wuwz@live.com).

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
package examples;

import javax.servlet.http.HttpServlet;

/**
 * @author wuwz
 */
// @WebServlet("/example")
public class ExampleServlet extends HttpServlet {
//
//	private static final long serialVersionUID = -8791212010764446339L;
//
//	@Override
//	protected void service(HttpServletRequest request, HttpServletResponse response)
//			throws ServletException, IOException {
//		request.setCharacterEncoding("UTF-8");
//		response.setCharacterEncoding("UTF-8");
//
//		String t = request.getParameter("t");
//		if ("list".equals(t)) {
//			toListPage(request, response);
//		}
//		// 导出
//		else if ("export".equals(t)) {
//
//			ExcelKit.$ExportRange(User.class, response)
//                    .setMaxSheetRecords(5)
//                    .toExcel(Db.getUsers(), "用户信息");
//		}
//		// 导入
//		else if ("import".equals(t)) {
//
//			importExcelFile(request, response);
//		}
//	}
//
//	private void importExcelFile(HttpServletRequest request, HttpServletResponse response) throws IOException {
//		PrintWriter writer = response.getWriter();
//		if (!ServletFileUpload.isMultipartContent(request)) {
//		    writer.println("Error: enctype!=multipart/form-data");
//		    writer.flush();
//		    return;
//		}
//
//		DiskFileItemFactory factory = new DiskFileItemFactory();
//		factory.setSizeThreshold(1024 * 1024 * 3);
//		factory.setRepository(new File(System.getProperty("java.io.tmpdir")));
//
//		ServletFileUpload upload = new ServletFileUpload(factory);
//		upload.setFileSizeMax(1024 * 1024 * 40);
//		upload.setSizeMax(1024 * 1024 * 50);
//
//		String uploadPath = request.getServletContext().getRealPath("./") + File.separator + "upload";
//		File uploadDir = new File(uploadPath);
//		if (!uploadDir.exists()) {
//		    uploadDir.mkdir();
//		}
//
//		try {
//		    List<FileItem> formItems = upload.parseRequest(request);
//
//		    if (formItems != null && formItems.size() > 0) {
//		        for (FileItem item : formItems) {
//		            if (!item.isFormField()) {
//		                String fileName = new File(item.getName()).getName();
//		                String filePath = uploadPath + File.separator + fileName;
//		                File storeFile = new File(filePath);
//		                System.out.println(filePath);
//		                item.write(storeFile);
//
//		                // 执行excel文件导入
//		                ExcelKit.$Import().readExcel(storeFile, new ReadHandler() {
//
//							@Override
//							public void handler(int sheetIndex, int rowIndex, List<String> row) {
//								if(rowIndex != 0) { //排除第一行
//									User user = new User()
//											.setUid(Integer.parseInt(row.get(0)))
//											.setUsername(row.get(1))
//											.setPassword(row.get(2))
//											.setSex(Integer.parseInt(row.get(3)))
//											.setGradeId(Integer.parseInt(row.get(4)));
//									Db.addUser(user);
//								}
//
//							}
//						});
//
//
//		                if(storeFile.exists()) {
//		                	storeFile.delete();
//		                }
//		                toListPage(request, response);
//		            }
//		        }
//		    }
//		} catch (Exception ex) {
//			writer.println("Error: "+ex.getMessage());
//			writer.flush();
//		}
//	}
//
//	private void toListPage(HttpServletRequest request, HttpServletResponse response)
//			throws ServletException, IOException {
//		request.setAttribute("users", Db.getUsers());
//		request.getRequestDispatcher("/list.jsp").forward(request, response);
//	}

}
