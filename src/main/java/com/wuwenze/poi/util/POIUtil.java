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

import com.wuwenze.poi.config.Options;
import com.wuwenze.poi.exception.ExcelKitException;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

import javax.servlet.http.HttpServletResponse;

/**
 * @author wuwenze
 * @date 2018/5/1
 */
public class POIUtil {

    private POIUtil() {
    }

    private static final int mDefaultRowAccessWindowSize = 100;


    public static SXSSFWorkbook newSXSSFWorkbook(int rowAccessWindowSize) {
        return new SXSSFWorkbook(rowAccessWindowSize);
    }

    public static SXSSFWorkbook newSXSSFWorkbook() {
        return newSXSSFWorkbook(mDefaultRowAccessWindowSize);
    }

    public static SXSSFSheet newSXSSFSheet(SXSSFWorkbook wb, String sheetName) {
        return (SXSSFSheet) wb.createSheet(sheetName);
    }

    public static SXSSFRow newSXSSFRow(SXSSFSheet sheet, int index) {
        return (SXSSFRow) sheet.createRow(index);
    }

    public static SXSSFCell newSXSSFCell(SXSSFRow row, int index) {
        return (SXSSFCell) row.createCell(index);
    }

    public static void setColumnWidth(SXSSFSheet sheet, int index, Short width, String value) {
        boolean widthNotHaveConfig = (null == width || width == -1);
        if (widthNotHaveConfig && !ValidatorUtil.isEmpty(value)) {
            sheet.setColumnWidth(index, (short) (value.length() * 2048));
        } else {
            width = widthNotHaveConfig ? 200 : width;
            sheet.setColumnWidth(index, (short) (width * 35.7));
        }
    }

    public static void setColumnCellRange(SXSSFSheet sheet, Options options, int firstRow, int endRow,
                                          int firstCell, int endCell) {
        if (null != options) {
            String[] datasource = options.get();
            if (null != datasource && datasource.length > 0) {
                if (datasource.length > 100) {
                    throw new ExcelKitException("Options item too much.");
                }

                DataValidationHelper validationHelper = sheet.getDataValidationHelper();
                DataValidationConstraint explicitListConstraint = validationHelper.createExplicitListConstraint(datasource);
                CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCell, endCell);
                DataValidation validation = validationHelper.createValidation(explicitListConstraint, regions);
                validation.setSuppressDropDownArrow(true);
                validation.createErrorBox("提示", "请从下拉列表选取");
                validation.setShowErrorBox(true);
                sheet.addValidationData(validation);
            }
        }
    }

    public static void write(SXSSFWorkbook wb, OutputStream out) {
        try {
            if (null != out) {
                wb.write(out);
                out.flush();
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (null != out) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void download(SXSSFWorkbook wb, HttpServletResponse response, String filename) {
        try {
            OutputStream out = response.getOutputStream();
            response.setContentType(Const.XLSX_CONTENT_TYPE);
            response.setHeader(Const.XLSX_HEADER_KEY,
                    String.format(Const.XLSX_HEADER_VALUE_TEMPLATE, filename));
            write(wb, out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static Object convertByExp(Object propertyValue, String converterExp) throws Exception {
        try {
            String[] convertSource = converterExp.split(",");
            for (String item : convertSource) {
                String[] itemArray = item.split("=");
                if (itemArray[0].equals(propertyValue)) {
                    return itemArray[1];
                }
            }
        } catch (Exception e) {
            throw e;
        }
        return propertyValue;
    }

    /**
     * 计算两个单元格之间的单元格数目(同一行)
     *
     * @return int
     */
    public static int countNullCell(String ref, String ref2) {
        // excel2007最大行数是1048576，最大列数是16384，最后一列列名是XFD
        String xfd = ref.replaceAll("\\d+", "");
        String xfd_1 = ref2.replaceAll("\\d+", "");

        xfd = fillChar(xfd, 3, '@', true);
        xfd_1 = fillChar(xfd_1, 3, '@', true);

        char[] letter = xfd.toCharArray();
        char[] letter_1 = xfd_1.toCharArray();
        int res = (letter[0] - letter_1[0]) * 26 * 26 + (letter[1] - letter_1[1]) * 26 + (letter[2] - letter_1[2]);
        return res - 1;
    }

    private static String fillChar(String str, int len, char let, boolean isPre) {
        int len_1 = str.length();
        if (len_1 < len) {
            if (isPre) {
                for (int i = 0; i < (len - len_1); i++) {
                    str = let + str;
                }
            } else {
                for (int i = 0; i < (len - len_1); i++) {
                    str = str + let;
                }
            }
        }
        return str;
    }

    public static void checkExcelFile(File file) {
        String filename = null != file ? file.getAbsolutePath() : null;
        if (null == filename || !file.exists()) {
            throw new ExcelKitException("Excel file[" + filename + "] does not exist.");
        }
        if (!filename.endsWith(Const.XLSX_SUFFIX)) {
            throw new ExcelKitException("[" + filename + "]Only .xlsx formatted files are supported.");
        }
    }



}
