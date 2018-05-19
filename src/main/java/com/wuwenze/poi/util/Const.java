/*
 * Copyright (c) 2018, 吴汶泽 (wuwz@live.com).
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *    http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package com.wuwenze.poi.util;

/**
 * @author wuwenze
 * @date 2018/5/1
 */
public final class Const {

  private Const() {}
  public static final String ENCODING = "UTF-8";
  public static final String XLSX_SUFFIX = ".xlsx";
  public static final String XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  public static final String XLSX_HEADER_KEY = "Content-disposition";
  public static final String XLSX_HEADER_VALUE_TEMPLATE = "attachment; filename=%s";
  public static final String XLSX_DEFAULT_EMPTY_CELL_VALUE = "$EMPTY_CELL$";
  public static final Integer XLSX_DEFAULT_BEGIN_READ_ROW_INDEX = 1;
  public static final String SAX_PARSER_CLASS = "org.apache.xerces.parsers.SAXParser";
  public static final String SAX_C_ELEMENT = "c";
  public static final String SAX_R_ATTR = "r";
  public static final String SAX_T_ELEMENT = "t";
  public static final String SAX_S_ATTR_VALUE = "s";
  public static final String SAX_RID_PREFIX = "rId";
  public static final String SAX_ROW_ELEMENT = "row";
}
